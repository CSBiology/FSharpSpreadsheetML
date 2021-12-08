namespace FSharpSpreadsheetML

module internal Dictionary =

    let tryGetValue k (dict:System.Collections.Generic.Dictionary<'K,'V>) = 
        let b,v = dict.TryGetValue(k)
        // Only get value if 
        if b 
        then 
            Some v
        else 
            None

    let length (dict:System.Collections.Generic.Dictionary<'K,'V>) = 
        dict.Count

type FieldMap<'T> =
    {
        CellTransformers : ('T -> XCell -> XCell) list
        HeaderTransformers : ('T -> XCell -> XCell) list
        ColumnWidth : float option
        RowHeight : ('T -> float option) option
        AdjustToContents: bool
    }
    with
        static member empty<'T>() = {
            CellTransformers = []
            HeaderTransformers = []
            ColumnWidth = None
            RowHeight = None
            AdjustToContents = false
        }

        static member create<'T>(mapRow: 'T -> XCell -> XCell) =
            let empty = FieldMap<'T>.empty()
            { empty with CellTransformers = List.append empty.CellTransformers [mapRow] }

        member self.header(name: string) =
            let transformer _ (cell: XCell) = cell.SetValue(name)
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.header(mapHeader: 'T -> string) =
            let transformer (value : 'T) (cell: XCell) = cell.SetValue(mapHeader value)
            { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

        member self.adjustToContents() =
            { self with AdjustToContents = true }

type Excel() =

    static member field<'T>(map: 'T -> int) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> string) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> System.DateTime) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> bool) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> double) = FieldMap<'T>.create(fun row cell -> cell.SetValue(map row))
    static member field<'T>(map: 'T -> int option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> System.DateTime option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> bool option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> double option) = FieldMap<'T>.create(fun row cell -> cell.SetValue(Option.toNullable (map row)))
    static member field<'T>(map: 'T -> string option) = FieldMap<'T>.create(fun row cell ->
        match map row with
        | None -> cell
        | Some text -> cell.SetValue(text)
    )

    static member populate<'T>(sheet: XWorksheet, data: seq<'T>, fields: FieldMap<'T> list) : unit =
        let headerTransformerGroups = fields |> List.map (fun field -> field.HeaderTransformers)
        let noHeadersAvailable =
            headerTransformerGroups
            |> List.concat
            |> List.isEmpty

        let headersAvailable = not noHeadersAvailable
       
        let headers : System.Collections.Generic.Dictionary<string,int> = System.Collections.Generic.Dictionary()

        //if headersAvailable then
        //    for (headerIndex, headerTransformers) in List.indexed headerTransformerGroups do
        //        let activeHeaderCell = sheet.Row(1).Cell(headerIndex + 1)
        //        for header in headerTransformers do ignore (header (Seq.head data) activeHeaderCell)

        for (rowIndex, row) in Seq.indexed data do
            let startRowIndex = if headersAvailable then 2 else 1
            let activeRow = sheet.Row(rowIndex + startRowIndex)
            for field in fields do

                let headerCell = XCell()
                for header in field.HeaderTransformers do ignore (header row headerCell)
                
                let index = 
                    match Dictionary.tryGetValue (headerCell.GetValue()) headers with
                    | Some int -> int
                    | None ->
                        let v = headerCell.GetValue()
                        let i = headers.Count + 1
                        headers.Add(v,i)
                        sheet.Row(1).Cell(i).CopyFrom(headerCell) |> ignore
                        i

                let activeCell = activeRow.Cell(index)
               
                for transformer in field.CellTransformers do
                    ignore (transformer row activeCell)

                //if field.AdjustToContents then
                //    let currentColumn = activeCell.WorksheetColumn()
                //    currentColumn.AdjustToContents() |> ignore
                //    activeRow.AdjustToContents() |> ignore

                //match field.ColumnWidth with
                //| Some givenWidth ->
                //    let currentColumn = activeCell.WorksheetColumn()
                //    currentColumn.Width <- givenWidth
                //| None -> ()

                //match field.RowHeight with
                //| Some givenHeightFn ->
                //    match givenHeightFn row with
                //    | Some givenHeight ->
                //        activeRow.Height <- givenHeight
                //    | None ->
                //        ()
                //| None ->
                //    ()


    static member workbookToBytes(workbook: XWorkbook) =
        use memoryStream = new System.IO.MemoryStream()
        workbook.SaveAs(memoryStream)
        memoryStream.ToArray()

    static member createFrom(name: string, data: seq<'T>, fields: FieldMap<'T> list) : byte[] =
        use workbook = new XWorkbook()
        let sheet = workbook.AddWorksheet(name)
        Excel.populate(sheet, data, fields)
        Excel.workbookToBytes(workbook)

    static member createFrom(workbook: XWorkbook) =
        use memoryStream = new System.IO.MemoryStream()
        workbook.SaveAs(memoryStream)
        memoryStream.ToArray()

    static member createFrom(data: seq<'T>, fields: FieldMap<'T> list) : byte[] =
        Excel.createFrom("Sheet1", data, fields)