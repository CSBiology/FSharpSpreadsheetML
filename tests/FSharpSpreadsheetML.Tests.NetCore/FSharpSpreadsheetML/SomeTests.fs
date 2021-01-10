module SomeTests

open FSharpSpreadsheetML
open Expecto


[<Tests>]
let testStuff = 
    testList "AminoAcids" [
        testCase "example" (fun () -> 
            let testSymbols = ['a','b','c']
            Expect.sequenceEqual
                testSymbols
                ['a','b','c']
                "wrooong"
        )
    ]