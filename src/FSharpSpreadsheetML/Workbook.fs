namespace FSharpSpreadsheetML

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging

/// Functions for manipulating workbooks. (Unmanaged: changing the sheets does not alter the associated worksheets which store the data)
module Workbook =

    /// Creates an empty workbook.
    let empty () = new Workbook()
    
    /// Gets the workbook of the workbookPart.
    let get (workbookPart : WorkbookPart) = workbookPart.Workbook 

    /// Sets the workbook of the workbookPart.
    let set (workbook : Workbook) (workbookPart : WorkbookPart) = 
        workbookPart.Workbook <- workbook
        workbookPart

    /// Sets an empty workbook.
    let init (workbookPart : WorkbookPart) = 
        set (Workbook()) workbookPart

    /// Returns the existing or a newly created workbook associated with the workbookpart.
    let getOrInit (workbookPart : WorkbookPart) =
        if workbookPart.Workbook <> null then
            get workbookPart
        else 
            workbookPart
            |> init
            |> get

    //// Adds sheet to workbook
    //let addSheet (sheet : Sheet) (workbook:Workbook) =
    //    let sheets = Sheet.Sheets.getOrInit workbook
    //    Sheet.Sheets.addSheet sheet sheets |> ignore
    //    workbook

/// Functions for working with WorkbookParts.
module WorkbookPart = 

    /// Add a worksheetPart to the workbookPart.
    let addWorksheetPart (worksheetPart : WorksheetPart) (workbookPart : WorkbookPart) = 
        workbookPart.AddPart(worksheetPart)

    /// Add an empty worksheetPart to the workbookPart.
    let initWorksheetPart (workbookPart : WorkbookPart) = workbookPart.AddNewPart<WorksheetPart>()

    /// Get the worksheetParts of the workbookPart.
    let getWorkSheetParts (workbookPart : WorkbookPart) = workbookPart.WorksheetParts

    /// Returns true if the workbookpart contains at least one worksheetPart.
    let containsWorkSheetParts (workbookPart : WorkbookPart) = workbookPart.GetPartsOfType<WorksheetPart>() |> Seq.length |> (<>) 0

    /// Gets the worksheetPart of the workbookPart with the given id.
    let getWorksheetPartById (id : string) (workbookPart : WorkbookPart) = workbookPart.GetPartById(id) :?> WorksheetPart 

    /// If the workbookpart contains the worksheetpart with the given id, returns it. Else returns None.
    let tryGetWorksheetPartById (id : string) (workbookPart : WorkbookPart) = 
        try workbookPart.GetPartById(id) :?> WorksheetPart  |> Some with
        | _ -> None

    /// Gets the ID of the worksheetPart of the workbookPart.
    let getWorksheetPartID (worksheetPart : WorksheetPart) (workbookPart : WorkbookPart) = workbookPart.GetIdOfPart worksheetPart
    //let addworkSheet (workbookPart:WorkbookPart) (worksheet : Worksheet) = 
    //    let emptySheet = (addNewWorksheetPart workbookPart)
    //    emptySheet.Worksheet <- worksheet

    /// Gets the sharedStringTablePart.
    let getSharedStringTablePart (workbookPart : WorkbookPart) = workbookPart.SharedStringTablePart
    
    /// Sets an empty sharedStringTablePart.
    let initSharedStringTablePart (workbookPart : WorkbookPart) = 
        workbookPart.AddNewPart<SharedStringTablePart>() |> ignore
        workbookPart

    /// Returns true if the workbookPart contains a sharedStringTablePart.
    let containsSharedStringTablePart (workbookPart : WorkbookPart) = workbookPart.SharedStringTablePart <> null

    /// Returns the existing or a newly created sharedStringTablePart associated with the workbookPart.
    let getOrInitSharedStringTablePart (workbookPart : WorkbookPart) =
        if containsSharedStringTablePart workbookPart then
            getSharedStringTablePart workbookPart
        else 
            initSharedStringTablePart workbookPart
            |> getSharedStringTablePart
    
    /// Returns the sharedStringTable of a workbookPart.
    let getSharedStringTable (workbookPart : WorkbookPart) =
        workbookPart 
        |> getSharedStringTablePart 
        |> SharedStringTable.get

    /// Returns the data of the first sheet of the given workbookPart.
    let getDataOfFirstSheet (workbookPart : WorkbookPart) = 
        workbookPart
        |> getWorkSheetParts
        |> Seq.head
        |> Worksheet.get
        |> Worksheet.getSheetData

    /// Appends a new sheet with the given sheet data to the MS Excel document.
    // to-do: guard if sheet of name already exists
    let appendSheet (sheetName : string) (data : SheetData) (workbookPart : WorkbookPart) =

        let workbook = Workbook.getOrInit  workbookPart

        let worksheetPart = initWorksheetPart workbookPart

        Worksheet.getOrInit worksheetPart
        |> Worksheet.addSheetData data
        |> ignore
        
        let sheets = Sheet.Sheets.getOrInit workbook
        let id = getWorksheetPartID worksheetPart workbookPart
        let sheetID = 
            sheets |> Sheet.Sheets.getSheets |> Seq.map Sheet.getSheetID
            |> fun s -> 
                if Seq.length s = 0 then 1u
                else s |> Seq.max |> (+) 1ul

        let sheet = Sheet.create id sheetName sheetID

        sheets.AppendChild(sheet) |> ignore
        workbookPart
