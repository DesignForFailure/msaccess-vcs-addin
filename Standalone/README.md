# Standalone Export/Import Modules

Two plain VBA modules that give you the core of the MSAccess VCS add-in without
installing it:

| Module | Purpose |
|---|---|
| `modVCSExport.bas` | Exports the current database to text source files in the add-in's format (export format 5.0.0). |
| `modVCSImport.bas` | Builds a new database from those files, or imports them into an existing database. |

Neither module needs extra VBA references. Everything is late bound, so they work in
any `.accdb`/`.mdb` with the default references, in 32-bit or 64-bit Access.

## Installing

1. Open the database in Access and press **Alt+F11** to open the VBA editor.
2. Choose **File > Import File...** and import `modVCSExport.bas` and/or
   `modVCSImport.bas`. Both modules can live in the same database.
3. Choose **Debug > Compile** to confirm that everything compiles.

The modules never export themselves, and the importer never replaces them while
they are running.

## Exporting

Run these in the Immediate window (**Ctrl+G**):

```vba
ExportSource                              ' -> <database file>.src\  (e.g. MyApp.accdb.src\)
ExportSource "C:\Work\MyApp.src"          ' export to a specific folder
ExportSource , "*"                        ' also save the data of every local table
ExportSource , "tblStates;tblColors"      ' also save the data of these tables
ExportObject acForm, "frmMain"            ' export one object (form, report, macro, module, query, table)
```

Tables selected for data export are remembered in `vcs-options.json` under
`TablesToExportData`, exactly like the add-in. They are saved as tab-delimited
`tables\*.txt` files. Set a table's `"Format"` to `"XML Format"` to use XML instead.

Each export also writes a log file to `<export folder>\logs\`.

## Importing / building

```vba
BuildFromSource "C:\Work\MyApp.accdb.src"
```
This creates a brand-new database next to the source folder, using the original
file name, in a separate Access window. It works like the add-in's
**Build From Source**.

- An existing file with that name is first renamed to `MyApp_VCSBackup.accdb`.
- Pass a second argument to choose the output file:
  `BuildFromSource "C:\Work\MyApp.accdb.src", "C:\Work\MyApp_new.accdb"`.
- Run it from any other database, such as a small "tools" database that holds
  `modVCSImport`.

```vba
MergeFromSource "C:\Work\MyApp.accdb.src"
```
This imports every source file into the **current** database.

- Forms, reports, queries, macros and modules are replaced.
- Existing local tables and their data are kept.
- New tables are created and filled with any exported data.
- `MergeFromSource , True` also replaces existing tables. Their data is lost unless
  the source contains data for them.
- Objects that exist in the database but not in the source are left alone.

```vba
ImportObject "C:\Work\MyApp.accdb.src\forms\frmMain.form"
```
This imports a single object into the current database. You can pass any of the
object's files (`.form`, `.cls`, `.json`, `.sql`, `.bas`, `tbldefs\*.xml`, ...).

- You are asked to confirm before a table that contains records is replaced.

After building or importing, open the VBA editor and run **Debug > Compile**.
Build and merge logs are written to `<source folder>\logs\`.

Start `MergeFromSource` and `ImportObject` from the Immediate window, not from a form
button or from another module that the import might replace. Both functions replace
VBA modules in the database that is running the code. Modules are imported last, and
missing references are added at the very end, to keep this safe. `BuildFromSource`
works in a separate copy of Access, so this does not apply to it.

## Typical workflow for editing with an AI assistant or a text editor

1. Run `ExportSource` in your database.
2. Edit the files in the `.src` folder (or have them edited), for example
   `forms\frmMain.form` / `forms\frmMain.cls`, `queries\qryX.sql`, or
   `modules\modY.bas`.
3. Bring the changes back into Access:
   - with `BuildFromSource` for a fresh database, or
   - with `ImportObject` / `MergeFromSource` for the database you are working in.

## File format

The folder layout and file formats match the add-in. Folders are `modules`,
`forms`, `reports`, `macros`, `queries`, `tbldefs`, `tables`, `relations`,
`tdmacros`, `themes`, `images`, `imexspecs` and `savedspecs`. Root-level files are
`vcs-options.json`, `project.json`, `vbe-project.json`, `vbe-references.json`,
`dbs-properties.json`, `proj-properties.json` and `documents.json`.

The details also match:
- UTF-8 with a BOM, and CRLF line endings
- `%`-encoded file names
- `'@Folder("A.B")` subfolders for modules, forms and reports
- form/report code split into a `.cls` file
- printer settings, descriptions and the hidden flag in companion `.json` files
- the same sanitizing rules

`Version Control.accda.src/AGENTS.md` in this repository describes the format in
detail. Copy it into your export folder if an AI assistant will be editing the files.

Differences from the add-in. All of these stay import-compatible in both directions:

- **Queries:** SQL is written as Access stores it, not reformatted. The query `.json`
  has no `DesignLayout`. On import, queries are created through DAO from the `.sql`
  file, so a query opens in SQL view until you save it again in Design View.
- **Conditional formatting:** it stays inline in the `.form`/`.report` file, which is
  the add-in's `DecodeConditionalFormatting = false` setting. The importer cannot
  rebuild conditional formatting that the add-in decoded into JSON.
- **Connection strings:** they are kept in the source files
  (`UseEnvForConnections = Never`), with passwords removed. `env:` references from
  add-in exports are resolved from a `.env` file in the source folder.
- **Not supported:** command bars (`menus`), navigation pane groups, VBE UserForms
  (`vbeforms`), ADP projects and external schema exports.
- **Non-ANSI characters in VBA:** characters that the system code page cannot
  represent are replaced when modules are imported. This is the same VBA editor
  limitation the add-in has, but here the log warns about it.
