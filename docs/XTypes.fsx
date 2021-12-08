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

let wb = new XWorkbook()

let sheet = wb.AddWorksheet("MySheet")

let row = sheet.Row(5)

let cell = row.Cell(10)

cell.Value <- "10"


let excelFile = Excel.createFrom(wb)
System.IO.File.WriteAllBytes(System.IO.Path.Combine(source,"Users.xlsx"), excelFile)




type User = { Name: string; Age: int ; Job : string}

let data = [
    { Name = "Jane"; Age = 26 ; Job = "Bowling"}
    { Name = "John"; Age = 25 ; Job = "Dude"}
    { Name = "Igor"; Age = 23 ; Job = "Lellek"}
]

let fields = [
    Excel.field(fun user -> user.Name)
    Excel.field(fun user -> user.Age)
    Excel.field(fun user -> user.Job).header ("Job")
]

let excelFile2 = Excel.createFrom(data,fields)

System.IO.File.WriteAllBytes(System.IO.Path.Combine(source,"Users2.xlsx"), excelFile2)
