(*** condition: prepare ***)
#r @"nuget: DocumentFormat.OpenXml"
#r "../bin/FSharpSpreadsheetML/netstandard2.0/FSharpSpreadsheetML.dll"
(*** condition: ipynb ***)
#if IPYNB
#r "nuget: Plotly.NET, {{fsdocs-package-version}}"
#r "nuget: Plotly.NET.Interactive, {{fsdocs-package-version}}"
#endif // IPYNB

let source = __SOURCE_DIRECTORY__

open FSharpSpreadsheetML

(**
# Table
[![Binder](https://mybinder.org/badge_logo.svg)](https://mybinder.org/v2/gh/plotly/Plotly.NET/gh-pages?filepath=Table.fsx.ipynb)

The table object itself just stores the name, the headers and the area in which the table lies. The values are stored in the sheetData object associated with the same worksheet as the table.

Therefore in order to work with tables, one should retrieve both the sheetData and the table. The Value retrieval functions ask for both.

## Accessing Tables

We try to access the table in the following excel sheet:

![excelTable](img/TestTables.png)


*)

let path = source + @"\content\files\tableTestFile.xlsx"

// Open excel document in no edit mode
let doc = Spreadsheet.fromFile path false

let sst = Spreadsheet.tryGetSharedStringTable doc

// Get the worksheet bound to the first excel sheet 
let worksheet = Spreadsheet.tryGetWorksheetPartBySheetIndex 0u doc |> Option.get

// Get the data of the worksheet
let sheetData = Worksheet.WorksheetPart.getSheetData worksheet

//List table names
Table.list worksheet |> Seq.map Table.getName

(*** include-it ***)

// Get the table by the table name retrieved in the list
let table = Table.tryGetByNameBy ((=) "Table1") worksheet |> Option.get


Table.getColumnHeaders table
(*** include-it ***)

// Get values of a single column
Table.tryGetColumnValuesByColumnHeader sst sheetData "Column3" table
(*** include-it ***)

// Get and combine the values of a key and a value column
Table.tryGetKeyValuesByColumnHeaders sst sheetData "Key" "Column2" "" table
(*** include-it ***)

(**
# Accessing Pseudo Tables

Tables not being implemented containing the actual values, but just a reference to the area the values are placed in can be kind of a headache to handle (Showcased by the code above being a little clumsy).

But this can also be used in a kind of cheesy way, interactively declaring the a part of a sheet to be a table and accessing it the same way we saw above. 

To showcase this, the sheet above contains values arranged as a table, but not actually being a table. We access them by selecting the area they are situated in:

*)

let topLeftCell = "G5"
let bottomRightCell = "I11"

// or alternatively using the cellreference functions

let topLeftCell' = CellReference.ofIndices 7u 5u
let bottomRightCell' = CellReference.ofIndices 9u 11u

let area = Table.Area.ofBoundaries topLeftCell bottomRightCell

(*** include-value: area ***)

// The sheetData needs to be given, as the table needs headers. For this the first row of the area is used.
let pseudoTable = Table.tryCreateWithExistingHeaders sst sheetData "PseudoSheet" area |> Option.get

// Happy Tabling

Table.getColumnHeaders pseudoTable
(*** include-it ***)

Table.tryGetIndexedColumnValuesByColumnHeader sst sheetData "Column2" pseudoTable
(*** include-it ***)
