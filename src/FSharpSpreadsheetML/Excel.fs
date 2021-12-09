namespace FSharpSpreadsheetML

open FSharpSpreadsheetML

module Excel = 
 
    module internal Dictionary =

        let seed = System.Random()

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
            Hash : string
        }
        with
            static member empty<'T>() = {
                CellTransformers = []
                HeaderTransformers = []
                ColumnWidth = None
                RowHeight = None
                AdjustToContents = false
                Hash = System.Guid.NewGuid().ToString()
            }

            static member create<'T>(mapRow: 'T -> XCell -> XCell) =
                let empty = FieldMap<'T>.empty()
                { empty with 
                    CellTransformers = List.append empty.CellTransformers [mapRow] 
                }

            member self.header(name: string) =
                let transformer _ (cell: XCell) = cell.SetValue(name)
                { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

            member self.header(mapHeader: 'T -> string) =
                let transformer (value : 'T) (cell: XCell) = cell.SetValue(mapHeader value)
                { self with HeaderTransformers = List.append self.HeaderTransformers [transformer] }

            member self.adjustToContents() =
                { self with AdjustToContents = true }

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

    type XWorksheet with
        
        static member populate<'T>(sheet: XWorksheet, data: seq<'T>, fields: FieldMap<'T> list) : unit =
            let headerTransformerGroups = fields |> List.map (fun field -> field.HeaderTransformers)
            let noHeadersAvailable =
                headerTransformerGroups
                |> List.concat
                |> List.isEmpty

            let headersAvailable = not noHeadersAvailable
       
            let headers : System.Collections.Generic.Dictionary<string,int> = System.Collections.Generic.Dictionary()

            for (rowIndex, row) in Seq.indexed data do
                let startRowIndex = if headersAvailable then 2 else 1
                let activeRow = sheet.Row(rowIndex + startRowIndex)
                for field in fields do

                    let headerCell = XCell()
                    for header in field.HeaderTransformers do ignore (header row headerCell)
                
                    let index = 
                        let hasHeader, headerString = 
                            if headerCell.GetValue() = "" then 
                                false, field.Hash 
                            else true, headerCell.GetValue()
                        printfn "%b,%s" hasHeader headerString
                        match Dictionary.tryGetValue (headerString) headers with
                        | Some int -> int
                        | None ->
                            let v = headerString
                            let i = headers.Count + 1
                            headers.Add(v,i)
                            if hasHeader then
                                sheet.Row(1).Cell(i).CopyFrom(headerCell) |> ignore
                            i
                    printfn "Index %i" index
                    let activeCell = activeRow.Cell(index)
                    printfn "Address %s" activeCell.Address.Address
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

            sheet.SortRows()

        static member createFrom(name: string, data: seq<'T>, fields: FieldMap<'T> list) (*: byte[]*) =          
            let sheet = XWorksheet(name)
            XWorksheet.populate(sheet, data, fields)
            sheet

        static member createFrom(data: seq<'T>, fields: FieldMap<'T> list) (*: byte[]*) =
            XWorksheet.createFrom("Sheet1", data, fields)
            
    type XWorkbook with

        static member populate<'T>(workbook: XWorkbook, name : string, data: seq<'T>, fields: FieldMap<'T> list) : unit =
            let sheet = workbook.AddWorksheet(name)
            XWorksheet.populate(sheet,data,fields)          

        static member createFrom(name: string, data: seq<'T>, fields: FieldMap<'T> list) (*: byte[]*) =
            let workbook = new XWorkbook()
            XWorkbook.populate(workbook, name, data, fields)
            workbook

        static member createFrom(data: seq<'T>, fields: FieldMap<'T> list) (*: byte[]*) =
            XWorkbook.createFrom("Sheet1", data, fields)

        static member createFrom(sheets : XWorksheet list)=
            let workbook = new XWorkbook()
            sheets
            |> List.iter (fun sheet -> workbook.AddWorksheet(sheet) |> ignore)
            workbook

        static member ToBytes(workbook: XWorkbook) =
            use memoryStream = new System.IO.MemoryStream()
            workbook.SaveAs(memoryStream)
            memoryStream.ToArray()

        static member ToFile(path,workbook: XWorkbook) =
            XWorkbook.ToBytes workbook
            |> fun bytes -> System.IO.File.WriteAllBytes (path, bytes)