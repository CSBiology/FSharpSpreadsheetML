module FSharpSpreadsheetML.Tests.NetCore

open Expecto

[<EntryPoint>]
let main argv =

    //FSharpSpreadsheetML core tests
    Tests.runTestsWithCLIArgs [] argv SomeTests.testStuff         |> ignore
    0