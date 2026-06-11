# MDB to DB Converter

This tool converts Microsoft Access `.mdb` or `.accdb` files into SQLite `.db`
files. The output database can be connected to from LibreOffice Base.

## Setup

1. Install Python 3.
2. Install the Microsoft Access Database Engine so Windows has an Access ODBC
   driver.
3. Install the Python dependency:

```powershell
py -3 -m pip install -r requirements.txt
```

You can check that the Access driver is visible with:

```powershell
py -3 mdb_to_db.py --list-drivers
```

## Use the GUI

Double-click `run_converter.bat`, choose your `.mdb` file, choose where to save
the `.db` file, then click **Convert**.

## Use the command line

```powershell
py -3 mdb_to_db.py "C:\path\source.mdb" "C:\path\output.db" --overwrite
```

If you omit the output path, the converter writes a `.db` file next to the
source file:

```powershell
py -3 mdb_to_db.py "C:\path\source.mdb" --overwrite
```

Password-protected databases can be converted with:

```powershell
py -3 mdb_to_db.py "C:\path\source.mdb" --password "your-password" --overwrite
```

## Open in LibreOffice

Open LibreOffice Base, choose to connect to an existing database, select SQLite,
then pick the generated `.db` file. If your LibreOffice installation does not
show SQLite as an option, connect through an installed SQLite ODBC driver.

## Notes

- Tables, columns, primary keys, and row data are copied.
- Access-specific forms, reports, macros, and VBA are not copied because SQLite
  stores data, not Access application objects.
- Date/time values are stored as ISO text so LibreOffice and SQLite tools can
  read them cleanly.
