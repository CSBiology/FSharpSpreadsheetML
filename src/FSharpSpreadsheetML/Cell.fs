namespace FSharpSpreadsheetML

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet


/// Functions for creating and manipulating Cells.
module Cell =


    /// Functions for manipulating CellValues.
    module CellValue = 

        /// Creates an empty CellValue.
        let empty() = CellValue()

        /// Create a new cellValue containing the given string.
        let create (value : string) = CellValue(value)

        /// Returns the value stored inside the CellValue.
        let getValue (cellValue : CellValue) = cellValue.Text

        /// Sets the value inside the CellValue.
        let setValue (value : string) (cellValue : CellValue) =  cellValue.Text <- value


    /// Creates an empty cell.
    let empty () = Cell()

    /// Returns the proper CellValues case for the given value.
    let inferCellValue (value : 'T) = 
        let value = box value
        match value with
        | :? char as c -> CellValues.String,c.ToString()
        | :? string as s -> CellValues.String,s.ToString()
        | :? bool as c -> CellValues.Boolean,c.ToString()
        | :? byte as i -> CellValues.Number,i.ToString()
        | :? sbyte as i -> CellValues.Number,i.ToString()
        | :? int as i -> CellValues.Number,i.ToString()
        | :? int16 as i -> CellValues.Number,i.ToString()
        | :? int64 as i -> CellValues.Number,i.ToString()
        | :? uint as i -> CellValues.Number,i.ToString()
        | :? uint16 as i -> CellValues.Number,i.ToString()
        | :? uint64 as i -> CellValues.Number,i.ToString()
        | :? single as i -> CellValues.Number,i.ToString()
        | :? float as i -> CellValues.Number,i.ToString()
        | :? decimal as i -> CellValues.Number,i.ToString()
        | :? System.DateTime as d -> CellValues.Date,d.Date.ToString()
        | _ ->  CellValues.String,value.ToString()

    /// Creates a cell from a CellValues type case, a "A1" style reference, and a CellValue containing the value string.
    let create (dataType : CellValues) (reference : string) (value : CellValue) = 
        Cell(CellReference = StringValue.FromString reference, DataType = EnumValue(dataType), CellValue = value)


    /// Create a cell using a shared string table, also returns the updated shared string table.
    let fromValue (sharedStringTable : SharedStringTable Option) columnIndex rowIndex (value : 'T) = 
        let value = box value
        match value with
        | :? string as s when sharedStringTable.IsSome-> 
            let sharedStringTable = sharedStringTable.Value
            let reference = CellReference.ofIndices columnIndex (rowIndex)
            match SharedStringTable.tryGetIndexByString s sharedStringTable with
            | Some i -> 
                i
                |> string
                |> CellValue.create
                |> create CellValues.SharedString reference
            | None ->
                let updatedSharedStringTable = 
                    sharedStringTable
                    |> SharedStringTable.SharedStringItem.add (SharedStringTable.SharedStringItem.create s) 

                updatedSharedStringTable
                |> SharedStringTable.count
                |> string
                |> CellValue.create
                |> create CellValues.SharedString reference 
        | _  -> 
           let valType,value = inferCellValue value
           let reference = CellReference.ofIndices columnIndex (rowIndex)
           create valType reference (CellValue.create value)

    /// Gets "A1"-style cell reference.
    let getReference (cell : Cell) = cell.CellReference.Value

    /// Sets "A1"-style cell reference.
    let setReference (reference) (cell : Cell) = 
        cell.CellReference <- StringValue.FromString reference
        cell

    /// Gets Some type if existent. Else returns None.
    let tryGetType (cell : Cell) = 
        if cell.DataType <> null then
            Some cell.DataType.Value
        else
            None
    
    /// Gets a Cell type.
    let getType (cell : Cell) = cell.DataType.Value

    /// Sets a Cell type.
    let setType (dataType : CellValues) (cell : Cell) = 
        cell.DataType <- EnumValue(dataType)
        cell

    /// Gets Some CellValue if cellValue is existent. Else returns None.
    let tryGetCellValue (cell : Cell) = 
        if cell.CellValue <> null then
            Some cell.CellValue
        else
            None

    /// Gets the CellValue.
    let getCellValue (cell : Cell) = cell.CellValue
    
    /// Maps a cell to the value string using a shared string table.
    let tryGetValue (sharedStringTable:SharedStringTable Option) (cell:Cell) =
        match cell |> tryGetType with
        | Some (CellValues.SharedString) when sharedStringTable.IsSome->
            let sharedStringTable = sharedStringTable.Value
            cell
            |> tryGetCellValue
            |> Option.map (
                CellValue.getValue 
                >> int
                >> fun i -> SharedStringTable.getText i sharedStringTable
                >> SharedStringTable.SharedStringItem.getText                   
            )
    
        | _ ->
            cell
            |> tryGetCellValue
            |> Option. map CellValue.getValue   



    /// Maps a Cell to the value string using a sharedStringTable.
    let getValue (sharedStringTable : SharedStringTable Option) (cell : Cell) =
        match cell |> tryGetType with
        | Some (CellValues.SharedString) when sharedStringTable.IsSome->
            let sharedStringTable = sharedStringTable.Value

            let sharedStringTableIndex = 
                cell
                |> getCellValue
                |> CellValue.getValue
                |> int

            sharedStringTable
            |> SharedStringTable.getText sharedStringTableIndex
            |> SharedStringTable.SharedStringItem.getText
        | _ ->
            cell
            |> getCellValue
            |> CellValue.getValue   

    /// Sets a CellValue.
    let setValue (value : CellValue) (cell : Cell) = 
        cell.CellValue <- value
        cell

    /// Includes a value from the sharedStringTable in Cell.CellValue.Text.
    let includeSharedStringValue (sharedStringTable:SharedStringTable) (cell:Cell) =
        if not (isNull cell.DataType) then  
            match cell |> tryGetType with
            | Some (CellValues.SharedString) ->
                let index = int cell.InnerText
                match sharedStringTable |> Seq.tryItem index with 
                | Some value -> 
                    cell.DataType <- EnumValue(CellValues.String)
                    cell.CellValue.Text <- value.InnerText
                | None ->
                    cell.CellValue.Text <- cell.InnerText
                cell  

            | _ -> cell
        else        
            cell.CellValue.Text <- cell.InnerText
            cell

type DataType = 
    | String
    | Boolean
    | Number
    | Date
    | Empty

    /// Returns the proper CellValues case for the given value.
    static member InferCellValue (value : 'T) = 
        let value = box value
        match value with
        | :? char as c -> DataType.String,c.ToString()
        | :? string as s -> DataType.String,s.ToString()
        | :? bool as c -> DataType.Boolean,c.ToString()
        | :? byte as i -> DataType.Number,i.ToString()
        | :? sbyte as i -> DataType.Number,i.ToString()
        | :? int as i -> DataType.Number,i.ToString()
        | :? int16 as i -> DataType.Number,i.ToString()
        | :? int64 as i -> DataType.Number,i.ToString()
        | :? uint as i -> DataType.Number,i.ToString()
        | :? uint16 as i -> DataType.Number,i.ToString()
        | :? uint64 as i -> DataType.Number,i.ToString()
        | :? single as i -> DataType.Number,i.ToString()
        | :? float as i -> DataType.Number,i.ToString()
        | :? decimal as i -> DataType.Number,i.ToString()
        | :? System.DateTime as d -> DataType.Date,d.Date.ToString()
        | _ ->  DataType.String,value.ToString()

// Type based on the type XLCell used in ClosedXml
type XCell (value : string, dataType : DataType)=
    
    let mutable _cellValue = value
    let mutable _dataType = dataType
    let mutable _comment  = raise (System.NotImplementedException())
    let mutable _hyperlink = raise (System.NotImplementedException())
    let mutable _richText = raise (System.NotImplementedException())
    let mutable _formulaA1 = raise (System.NotImplementedException())
    let mutable _formulaR1C1 = raise (System.NotImplementedException())

    let mutable _rowIndex : int = raise (System.NotImplementedException())
    let mutable _columnIndex : int = raise (System.NotImplementedException())

    new () = XCell ("", DataType.Empty)
    new (value : string) = XCell (value, DataType.String)
    new (value : int) = XCell (string value, DataType.Number)
    new (value : float) = XCell (string value, DataType.Number)

    member internal self.SharedStringId = raise (System.NotImplementedException())

    member self.Active = raise (System.NotImplementedException())
    
    /// <summary>Gets this cell's address, relative to the worksheet.</summary>
    /// <value>The cell's address.</value>
    member self.Address = XAddress
    
    /// <summary>
    /// Calculated value of cell formula. Is used for decreasing number of computations perfromed.
    /// May hold invalid value when <see cref="NeedsRecalculation"/> flag is True.
    /// </summary>
    member self.CachedValue = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns the current region. The current region is a range bounded by any combination of blank rows and blank columns
    /// </summary>
    /// <value>
    /// The current region.
    /// </value>
    member self.CurrentRegion = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets or sets the type of this cell's data.
    /// <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
    /// </summary>
    /// <value>
    /// The type of the cell's data.
    /// </value>
    /// <exception cref="ArgumentException"></exception>
    member self.DataType = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets or sets the cell's formula with A1 references.
    /// </summary>
    /// <value>The formula with A1 references.</value>
    member self.FormulaA1 = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets or sets the cell's formula with R1C1 references.
    /// </summary>
    /// <value>The formula with R1C1 references.</value>
    member self.FormulaR1C1 = raise (System.NotImplementedException())
    
    member self.FormulaReference = raise (System.NotImplementedException())
    
    member self.HasArrayFormula = raise (System.NotImplementedException())
    
    member self.HasComment = raise (System.NotImplementedException())
    
    member self.HasDataValidation = raise (System.NotImplementedException())
    
    member self.HasFormula = raise (System.NotImplementedException())
    
    member self.HasHyperlink = raise (System.NotImplementedException())
    
    member self.HasRichText = raise (System.NotImplementedException())
    
    member self.HasSparkline = raise (System.NotImplementedException())
    
    /// <summary>
    /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
    /// </summary>
    member self.NeedsRecalculation = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets or sets a value indicating whether this cell's text should be shared or not.
    /// </summary>
    /// <value>
    ///   If false the cell's text will not be shared and stored as an inline value.
    /// </value>
    member self.ShareString = raise (System.NotImplementedException())
    
    member self.Sparkline = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets or sets the cell's style.
    /// </summary>
    member self.Style = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets or sets the cell's value. To get or set a strongly typed value, use the GetValue&lt;T&gt; and SetValue methods.
    /// <para>ClosedXML will try to detect the data type through parsing. If it can't then the value will be left as a string.</para>
    /// <para>If the object is an IEnumerable, ClosedXML will copy the collection's data into a table starting from this cell.</para>
    /// <para>If the object is a range, ClosedXML will copy the range starting from this cell.</para>
    /// <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
    /// <para>If the value starts with a single quote, ClosedXML will assume the value is a text variable and will prefix the value with a single quote in Excel too.</para>
    /// </summary>
    /// <value>
    /// The object containing the value(s) to set.
    /// </value>
    member self.Value = raise (System.NotImplementedException())
    
    member self.Worksheet = raise (System.NotImplementedException())
    
    member self.AddConditionalFormat()  = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates a named range out of this cell.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <para>The default scope for the named range is Workbook.</para>
    /// </summary>
    /// <param name="rangeName">Name of the range.</param>
    member self.AddToNamed(rangeName)  = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates a named range out of this cell.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <param name="rangeName">Name of the range.</param>
    /// <param name="scope">The scope for the named range.</param>
    /// </summary>
    member self.AddToNamed(rangeName, scope) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates a named range out of this cell.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <param name="rangeName">Name of the range.</param>
    /// <param name="scope">The scope for the named range.</param>
    /// <param name="comment">The comments for the named range.</param>
    /// </summary>
    member self.AddToNamed(rangeName, scope, comment) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns this cell as an IXLRange.
    /// </summary>
    member self.AsRange()  = raise (System.NotImplementedException())
    
    member self.CellAbove() = raise (System.NotImplementedException())
    
    member self.CellAbove(step) = raise (System.NotImplementedException())
    
    member self.CellBelow() = raise (System.NotImplementedException())
    
    member self.CellBelow(step) = raise (System.NotImplementedException())
    
    member self.CellLeft() = raise (System.NotImplementedException())
    
    member self.CellLeft(step) = raise (System.NotImplementedException())
    
    member self.CellRight() = raise (System.NotImplementedException())
    
    member self.CellRight(step) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Clears the contents of this cell.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    member self.Clear(clearOptions(* = XLClearOptions.All*)) = raise (System.NotImplementedException())
    
    //member self.CopyFrom(member self.otherCell);
    
    member self.CopyFrom(otherCell) = raise (System.NotImplementedException())
    
    //member self.CopyTo(member self.target);
    
    member self.CopyTo(target) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates a new comment for the cell, replacing the existing one.
    /// </summary>
    member self.CreateComment() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates a new data validation rule for the cell, replacing the existing one.
    /// </summary>
    member self.CreateDataValidation() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates a new hyperlink replacing the existing one.
    /// </summary>
    member self.CreateHyperlink() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Replaces a value of the cell with a newly created rich text object.
    /// </summary>
    member self.CreateRichText() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Deletes the current cell and shifts the surrounding cells according to the shiftDeleteCells parameter.
    /// </summary>
    /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
    member self.Delete(shiftDeleteCells) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value converted to Boolean.
    /// <para>ClosedXML will try to covert the current value to Boolean.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to Boolean.</para>
    /// </summary>
    member self.GetBoolean() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns the comment for the cell or create a new instance if there is no comment on the cell.
    /// </summary>
    member self.GetComment() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns a data validation rule assigned to the cell, if any, or creates a new instance of data validation rule if no rule exists.
    /// </summary>
    member self.GetDataValidation() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value converted to DateTime.
    /// <para>ClosedXML will try to covert the current value to DateTime.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to DateTime.</para>
    /// </summary>
    member self.GetDateTime() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value converted to Double.
    /// <para>ClosedXML will try to covert the current value to Double.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to Double.</para>
    /// </summary>
    member self.GetDouble() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value formatted depending on the cell's data type and style.
    /// </summary>
    member self.GetFormattedString() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns a hyperlink for the cell, if any, or creates a new instance is there is no hyperlink.
    /// </summary>
    member self.GetHyperlink() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns the value of the cell if it formatted as a rich text.
    /// </summary>
    member self.GetRichText() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value converted to a String.
    /// </summary>
    member self.GetString() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value converted to TimeSpan.
    /// <para>ClosedXML will try to covert the current value to TimeSpan.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to TimeSpan.</para>
    /// </summary>
    member self.GetTimeSpan() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Gets the cell's value converted to the T type.
    /// <para>ClosedXML will try to covert the current value to the T type.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to the T type.</para>
    /// </summary>
    /// <typeparam name="T">The return type.</typeparam>
    /// <exception cref="ArgumentException"></exception>
    member self.GetValue<'T>() = raise (System.NotImplementedException())
    
    member self.GetValue() = _cellValue

    member self.InsertCellsAbove(numberOfRows) = raise (System.NotImplementedException())
    
    member self.InsertCellsAfter(numberOfColumns) = raise (System.NotImplementedException())
    
    member self.InsertCellsBefore(numberOfColumns) = raise (System.NotImplementedException())
    
    member self.InsertCellsBelow(numberOfRows)  = raise (System.NotImplementedException())
    
    /// <summary>
    /// Inserts the IEnumerable data elements and returns the range it occupies.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    member self.InsertData(data)  = raise (System.NotImplementedException())
    
    /// <summary>
    /// Inserts the IEnumerable data elements and returns the range it occupies.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
    /// <returns></returns>
    member self.InsertData(data, transpose) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the data of a data table.
    ///// </summary>
    ///// <param name="dataTable">The data table.</param>
    ///// <returns>The range occupied by the inserted data</returns>
    //member self.InsertData(dataTable) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Inserts the IEnumerable data elements as a table and returns it.
    /// <para>The new table will receive a generic name: Table#</para>
    /// </summary>
    /// <param name="data">The table data.</param>
    member self.InsertTable<'T>(data) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the IEnumerable data elements as a table and returns it.
    ///// <para>The new table will receive a generic name: Table#</para>
    ///// </summary>
    ///// <param name="data">The table data.</param>
    ///// <param name="createTable">
    ///// if set to <c>true</c> it will create an Excel table.
    ///// <para>if set to <c>false</c> the table will be created in memory.</para>
    ///// </param>
    //member self.InsertTable<'T>(data, createTable)  = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates an Excel table from the given IEnumerable data elements.
    /// </summary>
    /// <param name="data">The table data.</param>
    /// <param name="tableName">Name of the table.</param>
    member self.InsertTable<'T>(data, tableName) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Inserts the IEnumerable data elements as a table and returns it.
    /// </summary>
    /// <param name="data">The table data.</param>
    /// <param name="tableName">Name of the table.</param>
    /// <param name="createTable">
    /// if set to <c>true</c> it will create an Excel table.
    /// <para>if set to <c>false</c> the table will be created in memory.</para>
    /// </param>
    member self.InsertTable<'T>(data, tableName, createTable) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Inserts the DataTable data elements as a table and returns it.
    /// <para>The new table will receive a generic name: Table#</para>
    /// </summary>
    /// <param name="data">The table data.</param>
    member self.InsertTable(data) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the DataTable data elements as a table and returns it.
    ///// <para>The new table will receive a generic name: Table#</para>
    ///// </summary>
    ///// <param name="data">The table data.</param>
    ///// <param name="createTable">
    ///// if set to <c>true</c> it will create an Excel table.
    ///// <para>if set to <c>false</c> the table will be created in memory.</para>
    ///// </param>
    //member self.InsertTable(data, createTable) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Creates an Excel table from the given DataTable data elements.
    /// </summary>
    /// <param name="data">The table data.</param>
    /// <param name="tableName">Name of the table.</param>
    member self.InsertTable(data, tableName)  = raise (System.NotImplementedException())
    
    /// <summary>
    /// Inserts the DataTable data elements as a table and returns it.
    /// </summary>
    /// <param name="data">The table data.</param>
    /// <param name="tableName">Name of the table.</param>
    /// <param name="createTable">
    /// if set to <c>true</c> it will create an Excel table.
    /// <para>if set to <c>false</c> the table will be created in memory.</para>
    /// </param>
    member self.InsertTable(data, tableName, createTable) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Invalidate <see cref="CachedValue"/> so the formula will be re-evaluated next time <see cref="Value"/> is accessed.
    /// If cell does not contain formula nothing happens.
    /// </summary>
    member self.InvalidateFormula() = raise (System.NotImplementedException())
    
    member self.IsEmpty() = raise (System.NotImplementedException())
    
    [<System.Obsolete("Use the overload with XLCellsUsedOptions")>]
    member self.IsEmpty(includeFormats) = raise (System.NotImplementedException())
    
    //member self.IsEmpty(options) = raise (System.NotImplementedException())
    
    member self.IsMerged() = raise (System.NotImplementedException())
    
    member self.MergedRange() = raise (System.NotImplementedException())
    
    member self.Select() = raise (System.NotImplementedException())
    
    member self.SetActive(value(* = true*)) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Sets the type of this cell's data.
    /// <para>Changing the data type will cause ClosedXML to covert the current value to the new data type.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to the new data type.</para>
    /// </summary>
    /// <param name="dataType">Type of the data.</param>
    /// <returns></returns>
    member self.SetDataType(dataType) = raise (System.NotImplementedException())
    
    [<System.Obsolete("Use GetDataValidation to access the existing rule, or CreateDataValidation() to create a new one.")>]
    member self.SetDataValidation() = raise (System.NotImplementedException())
    
    member self.SetFormulaA1(formula) = raise (System.NotImplementedException())
    
    member self.SetFormulaR1C1(formula) = raise (System.NotImplementedException())
    
    member self.SetHyperlink(hyperlink) = raise (System.NotImplementedException())
    
    /// <summary>
    /// Sets the cell's value.
    /// <para>If the object is an IEnumerable ClosedXML will copy the collection's data into a table starting from this cell.</para>
    /// <para>If the object is a range ClosedXML will copy the range starting from this cell.</para>
    /// <para>Setting the value to an object (not IEnumerable/range) will call the object's ToString() method.</para>
    /// <para>ClosedXML will try to translate it to the corresponding type, if it can't then the value will be left as a string.</para>
    /// </summary>
    /// <value>
    /// The object containing the value(s) to set.
    /// </value>
    member self.SetValue<'T>(value) = 
        let t,v = DataType.InferCellValue value
        _dataType <- t
        _cellValue <- v
        self

    member self.TableCellType() = raise (System.NotImplementedException())
    
    /// <summary>
    /// Returns a string that represents the current state of the cell according to the format.
    /// </summary>
    /// <param name="format">A: address, F: formula, NF: number format, BG: background color, FG: foreground color, V: formatted value</param>
    /// <returns></returns>
    member self.ToString(format) = raise (System.NotImplementedException())
    
    member self.TryGetValue<'T>(value) = raise (System.NotImplementedException())
    
    member self.WorksheetColumn() = raise (System.NotImplementedException())
    
    member self.WorksheetRow() = raise (System.NotImplementedException())