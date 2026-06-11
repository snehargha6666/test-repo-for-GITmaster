from __future__ import annotations

import argparse
import datetime as dt
import decimal
import os
from pathlib import Path
import queue
import sqlite3
import sys
import threading
from dataclasses import dataclass
from typing import Any, Callable, Iterable


ProgressCallback = Callable[[str], None]


ACCESS_DRIVER_PREFERENCE = (
    "Microsoft Access Driver (*.mdb, *.accdb)",
    "Microsoft Access Driver (*.mdb)",
)


class ConversionError(RuntimeError):
    """Raised for user-fixable conversion problems."""


@dataclass(frozen=True)
class ColumnInfo:
    name: str
    access_type: str
    sqlite_type: str


@dataclass(frozen=True)
class TableResult:
    name: str
    rows: int


def load_pyodbc():
    try:
        import pyodbc  # type: ignore
    except ImportError as exc:
        raise ConversionError(
            "The 'pyodbc' package is required. Install it with:\n"
            "  py -3 -m pip install -r requirements.txt\n"
            "or:\n"
            "  python -m pip install -r requirements.txt"
        ) from exc
    return pyodbc


def find_access_driver(pyodbc_module: Any, requested_driver: str | None = None) -> str:
    if requested_driver:
        return requested_driver

    drivers = list(pyodbc_module.drivers())
    access_drivers = [
        driver
        for driver in drivers
        if "access" in driver.lower() and ".mdb" in driver.lower()
    ]

    for preferred in ACCESS_DRIVER_PREFERENCE:
        if preferred in access_drivers:
            return preferred

    if access_drivers:
        return access_drivers[-1]

    raise ConversionError(
        "No Microsoft Access ODBC driver was found on this computer.\n\n"
        "Install the Microsoft Access Database Engine, then run this converter again.\n"
        "You can check installed ODBC drivers with:\n"
        "  python mdb_to_db.py --list-drivers"
    )


def quote_access_identifier(identifier: str) -> str:
    return "[" + identifier.replace("]", "]]") + "]"


def quote_sqlite_identifier(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def sqlite_type_for_access(access_type: str, data_type: int | None) -> str:
    access_type_upper = (access_type or "").upper()

    integer_names = {
        "AUTOINCREMENT",
        "BYTE",
        "COUNTER",
        "INTEGER",
        "LONG",
        "LONG INTEGER",
        "SHORT",
        "SMALLINT",
        "TINYINT",
    }
    boolean_names = {"BIT", "BOOLEAN", "LOGICAL", "YESNO", "YES/NO"}
    real_names = {"DOUBLE", "FLOAT", "REAL", "SINGLE"}
    numeric_names = {"CURRENCY", "DECIMAL", "MONEY", "NUMERIC"}
    date_names = {"DATE", "DATETIME", "TIME", "TIMESTAMP"}
    blob_names = {
        "BINARY",
        "GENERAL",
        "IMAGE",
        "LONGBINARY",
        "LONGVARBINARY",
        "OLEOBJECT",
        "VARBINARY",
    }

    if access_type_upper in integer_names or access_type_upper in boolean_names:
        return "INTEGER"
    if access_type_upper in real_names:
        return "REAL"
    if access_type_upper in numeric_names:
        return "NUMERIC"
    if access_type_upper in date_names:
        return "TEXT"
    if access_type_upper in blob_names:
        return "BLOB"

    # ODBC SQL type ids, used when an Access driver reports a localized type name.
    if data_type in {-7, -6, 4, 5}:  # BIT, TINYINT, INTEGER, SMALLINT
        return "INTEGER"
    if data_type in {2, 3}:  # NUMERIC, DECIMAL
        return "NUMERIC"
    if data_type in {6, 7, 8}:  # FLOAT, REAL, DOUBLE
        return "REAL"
    if data_type in {9, 10, 11, 91, 92, 93}:  # DATE/TIME/TIMESTAMP variants
        return "TEXT"
    if data_type in {-4, -3, -2}:  # binary variants
        return "BLOB"

    return "TEXT"


def connect_access(
    mdb_path: Path,
    driver: str | None,
    password: str | None,
):
    pyodbc = load_pyodbc()
    selected_driver = find_access_driver(pyodbc, driver)
    connection_parts = [
        f"DRIVER={{{selected_driver}}}",
        f"DBQ={mdb_path}",
        "READONLY=TRUE",
    ]

    if password:
        connection_parts.append(f"PWD={password}")

    connection_string = ";".join(connection_parts) + ";"
    return pyodbc.connect(connection_string, autocommit=True), selected_driver


def list_user_tables(access_connection: Any, include_system: bool) -> list[str]:
    tables: list[str] = []
    cursor = access_connection.cursor()
    for row in cursor.tables():
        table_name = str(row.table_name)
        table_type = str(row.table_type).upper()
        is_system = table_name.upper().startswith("MSYS") or "SYSTEM" in table_type

        if table_type != "TABLE":
            continue
        if is_system and not include_system:
            continue

        tables.append(table_name)

    return sorted(dict.fromkeys(tables), key=str.lower)


def read_columns(access_connection: Any, table_name: str) -> list[ColumnInfo]:
    cursor = access_connection.cursor()
    columns: list[ColumnInfo] = []

    for row in cursor.columns(table=table_name):
        column_name = str(row.column_name)
        access_type = str(getattr(row, "type_name", "") or "")
        data_type = getattr(row, "data_type", None)
        columns.append(
            ColumnInfo(
                name=column_name,
                access_type=access_type,
                sqlite_type=sqlite_type_for_access(access_type, data_type),
            )
        )

    if not columns:
        raise ConversionError(f"Table '{table_name}' has no readable columns.")

    return columns


def read_primary_key_columns(access_connection: Any, table_name: str) -> list[str]:
    cursor = access_connection.cursor()
    try:
        rows = list(cursor.primaryKeys(table=table_name))
    except Exception:
        return []

    rows.sort(key=lambda row: getattr(row, "key_seq", 0))
    return [str(row.column_name) for row in rows]


def create_sqlite_table(
    sqlite_connection: sqlite3.Connection,
    table_name: str,
    columns: list[ColumnInfo],
    primary_key_columns: Iterable[str],
) -> None:
    column_names = {column.name for column in columns}
    valid_primary_keys = [name for name in primary_key_columns if name in column_names]

    definitions = [
        f"{quote_sqlite_identifier(column.name)} {column.sqlite_type}"
        for column in columns
    ]

    if valid_primary_keys:
        keys = ", ".join(quote_sqlite_identifier(name) for name in valid_primary_keys)
        definitions.append(f"PRIMARY KEY ({keys})")

    create_sql = (
        f"CREATE TABLE {quote_sqlite_identifier(table_name)} (\n"
        f"  {',\n  '.join(definitions)}\n"
        ")"
    )
    sqlite_connection.execute(create_sql)


def convert_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat(sep=" ") if isinstance(value, dt.datetime) else value.isoformat()
    if isinstance(value, decimal.Decimal):
        return str(value)
    if isinstance(value, memoryview):
        return bytes(value)
    if isinstance(value, bytearray):
        return bytes(value)
    return value


def copy_table_data(
    access_connection: Any,
    sqlite_connection: sqlite3.Connection,
    table_name: str,
    columns: list[ColumnInfo],
    batch_size: int,
) -> int:
    access_cursor = access_connection.cursor()
    access_column_list = ", ".join(quote_access_identifier(column.name) for column in columns)
    select_sql = f"SELECT {access_column_list} FROM {quote_access_identifier(table_name)}"
    access_cursor.execute(select_sql)

    sqlite_column_list = ", ".join(quote_sqlite_identifier(column.name) for column in columns)
    placeholders = ", ".join("?" for _ in columns)
    insert_sql = (
        f"INSERT INTO {quote_sqlite_identifier(table_name)} ({sqlite_column_list}) "
        f"VALUES ({placeholders})"
    )

    copied_rows = 0
    while True:
        rows = access_cursor.fetchmany(batch_size)
        if not rows:
            break

        sqlite_rows = [tuple(convert_value(value) for value in row) for row in rows]
        sqlite_connection.executemany(insert_sql, sqlite_rows)
        copied_rows += len(sqlite_rows)

    return copied_rows


def convert_mdb_to_sqlite(
    source: Path,
    destination: Path,
    *,
    overwrite: bool = False,
    include_system: bool = False,
    batch_size: int = 500,
    driver: str | None = None,
    password: str | None = None,
    progress: ProgressCallback | None = None,
) -> list[TableResult]:
    source = source.expanduser().resolve()
    destination = destination.expanduser().resolve()

    if not source.exists():
        raise ConversionError(f"Source file does not exist: {source}")
    if source.suffix.lower() not in {".mdb", ".accdb"}:
        raise ConversionError("Source file must be an .mdb or .accdb file.")
    if destination.exists() and not overwrite:
        raise ConversionError(
            f"Destination already exists: {destination}\n"
            "Use --overwrite or choose a different output path."
        )
    if batch_size < 1:
        raise ConversionError("Batch size must be at least 1.")

    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_destination = destination.with_name(destination.name + ".tmp")
    if temp_destination.exists():
        temp_destination.unlink()

    log = progress or (lambda message: None)
    access_connection = None
    sqlite_connection = None

    try:
        log(f"Opening Access database: {source}")
        access_connection, selected_driver = connect_access(source, driver, password)
        log(f"Using ODBC driver: {selected_driver}")

        tables = list_user_tables(access_connection, include_system)
        if not tables:
            raise ConversionError("No user tables were found in the Access database.")

        sqlite_connection = sqlite3.connect(temp_destination)
        sqlite_connection.execute("PRAGMA foreign_keys = OFF")
        sqlite_connection.execute("PRAGMA journal_mode = OFF")
        sqlite_connection.execute("PRAGMA synchronous = OFF")

        results: list[TableResult] = []
        for index, table_name in enumerate(tables, start=1):
            log(f"[{index}/{len(tables)}] Converting table: {table_name}")
            columns = read_columns(access_connection, table_name)
            primary_keys = read_primary_key_columns(access_connection, table_name)
            create_sqlite_table(sqlite_connection, table_name, columns, primary_keys)
            rows = copy_table_data(
                access_connection,
                sqlite_connection,
                table_name,
                columns,
                batch_size,
            )
            sqlite_connection.commit()
            results.append(TableResult(name=table_name, rows=rows))
            log(f"    {rows} row(s) copied")

        sqlite_connection.execute("PRAGMA optimize")
        sqlite_connection.commit()
        sqlite_connection.close()
        sqlite_connection = None

        os.replace(temp_destination, destination)
        log(f"Done: {destination}")
        return results
    except Exception:
        if temp_destination.exists():
            temp_destination.unlink()
        raise
    finally:
        if sqlite_connection is not None:
            sqlite_connection.close()
        if access_connection is not None:
            access_connection.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert Microsoft Access .mdb/.accdb files to SQLite .db files."
    )
    parser.add_argument(
        "source",
        nargs="?",
        help="Path to the .mdb or .accdb file. Opens the GUI if omitted.",
    )
    parser.add_argument(
        "destination",
        nargs="?",
        help="Path to the output .db file. Defaults to SOURCE with a .db extension.",
    )
    parser.add_argument("--overwrite", action="store_true", help="Replace the output file.")
    parser.add_argument(
        "--include-system",
        action="store_true",
        help="Also copy Access system tables such as MSysObjects.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=500,
        help="Rows to copy per insert batch. Default: 500.",
    )
    parser.add_argument("--driver", help="Use a specific Access ODBC driver name.")
    parser.add_argument("--password", help="Database password, if the .mdb is protected.")
    parser.add_argument("--gui", action="store_true", help="Open the file picker interface.")
    parser.add_argument(
        "--list-drivers",
        action="store_true",
        help="Print installed ODBC drivers and exit.",
    )
    return parser


def print_drivers() -> int:
    pyodbc = load_pyodbc()
    drivers = list(pyodbc.drivers())
    if not drivers:
        print("No ODBC drivers were found.")
        return 1

    print("Installed ODBC drivers:")
    for driver in drivers:
        marker = "  <-- Access candidate" if (
            "access" in driver.lower() and ".mdb" in driver.lower()
        ) else ""
        print(f"  {driver}{marker}")
    return 0


def run_cli(args: argparse.Namespace) -> int:
    if args.list_drivers:
        try:
            return print_drivers()
        except ConversionError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 2

    if args.gui or not args.source:
        return run_gui()

    source = Path(args.source)
    destination = Path(args.destination) if args.destination else source.with_suffix(".db")

    try:
        results = convert_mdb_to_sqlite(
            source,
            destination,
            overwrite=args.overwrite,
            include_system=args.include_system,
            batch_size=args.batch_size,
            driver=args.driver,
            password=args.password,
            progress=print,
        )
    except ConversionError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"Unexpected error: {exc}", file=sys.stderr)
        return 1

    print("\nConverted tables:")
    for result in results:
        print(f"  {result.name}: {result.rows} row(s)")
    return 0


def run_gui() -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except Exception as exc:
        print(f"Unable to open GUI: {exc}", file=sys.stderr)
        return 1

    root = tk.Tk()
    root.title("MDB to DB Converter")
    root.geometry("760x460")
    root.minsize(640, 380)

    source_var = tk.StringVar()
    destination_var = tk.StringVar()
    password_var = tk.StringVar()
    overwrite_var = tk.BooleanVar(value=True)
    status_var = tk.StringVar(value="Choose an Access database to convert.")
    log_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()

    main = ttk.Frame(root, padding=16)
    main.pack(fill="both", expand=True)
    main.columnconfigure(1, weight=1)
    main.rowconfigure(5, weight=1)

    def choose_source() -> None:
        selected = filedialog.askopenfilename(
            title="Choose Access database",
            filetypes=[
                ("Access databases", "*.mdb *.accdb"),
                ("MDB files", "*.mdb"),
                ("ACCDB files", "*.accdb"),
                ("All files", "*.*"),
            ],
        )
        if not selected:
            return

        source_var.set(selected)
        if not destination_var.get():
            destination_var.set(str(Path(selected).with_suffix(".db")))

    def choose_destination() -> None:
        selected = filedialog.asksaveasfilename(
            title="Save SQLite database",
            defaultextension=".db",
            filetypes=[
                ("SQLite database", "*.db"),
                ("SQLite database", "*.sqlite"),
                ("All files", "*.*"),
            ],
        )
        if selected:
            destination_var.set(selected)

    ttk.Label(main, text="Access file").grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(main, textvariable=source_var).grid(row=0, column=1, sticky="ew", padx=8, pady=(0, 8))
    ttk.Button(main, text="Browse", command=choose_source).grid(row=0, column=2, pady=(0, 8))

    ttk.Label(main, text="Output .db").grid(row=1, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(main, textvariable=destination_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(0, 8))
    ttk.Button(main, text="Browse", command=choose_destination).grid(row=1, column=2, pady=(0, 8))

    ttk.Label(main, text="Password").grid(row=2, column=0, sticky="w", pady=(0, 8))
    ttk.Entry(main, textvariable=password_var, show="*").grid(
        row=2,
        column=1,
        sticky="ew",
        padx=8,
        pady=(0, 8),
    )
    ttk.Checkbutton(main, text="Replace output file", variable=overwrite_var).grid(
        row=3,
        column=1,
        sticky="w",
        padx=8,
        pady=(0, 8),
    )

    status_label = ttk.Label(main, textvariable=status_var)
    status_label.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 8))

    log_text = tk.Text(main, height=12, wrap="word", state="disabled")
    log_text.grid(row=5, column=0, columnspan=3, sticky="nsew")

    scrollbar = ttk.Scrollbar(main, orient="vertical", command=log_text.yview)
    scrollbar.grid(row=5, column=3, sticky="ns")
    log_text.configure(yscrollcommand=scrollbar.set)

    button_row = ttk.Frame(main)
    button_row.grid(row=6, column=0, columnspan=3, sticky="e", pady=(12, 0))
    convert_button = ttk.Button(button_row, text="Convert")
    convert_button.pack(side="right")

    def append_log(message: str) -> None:
        log_text.configure(state="normal")
        log_text.insert("end", message + "\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    def set_busy(is_busy: bool) -> None:
        convert_button.configure(state="disabled" if is_busy else "normal")

    def worker() -> None:
        try:
            results = convert_mdb_to_sqlite(
                Path(source_var.get()),
                Path(destination_var.get()),
                overwrite=overwrite_var.get(),
                password=password_var.get() or None,
                progress=lambda message: log_queue.put(("log", message)),
            )
            total_rows = sum(result.rows for result in results)
            log_queue.put(("done", f"Converted {len(results)} table(s), {total_rows} row(s)."))
        except Exception as exc:
            log_queue.put(("error", str(exc)))

    def start_conversion() -> None:
        if not source_var.get():
            messagebox.showerror("Missing file", "Choose an .mdb or .accdb file first.")
            return
        if not destination_var.get():
            messagebox.showerror("Missing output", "Choose where to save the .db file.")
            return

        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")
        status_var.set("Converting...")
        set_busy(True)
        threading.Thread(target=worker, daemon=True).start()

    def poll_queue() -> None:
        try:
            while True:
                kind, message = log_queue.get_nowait()
                if kind == "log":
                    append_log(message)
                elif kind == "done":
                    append_log(message)
                    status_var.set("Conversion complete.")
                    set_busy(False)
                    messagebox.showinfo("Done", message)
                elif kind == "error":
                    append_log("Error: " + message)
                    status_var.set("Conversion failed.")
                    set_busy(False)
                    messagebox.showerror("Conversion failed", message)
        except queue.Empty:
            pass
        root.after(100, poll_queue)

    convert_button.configure(command=start_conversion)
    root.after(100, poll_queue)
    root.mainloop()
    return 0


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
