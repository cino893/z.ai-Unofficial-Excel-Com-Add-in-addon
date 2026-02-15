Attribute VB_Name = "modExcelSkills"
'==============================================================================
' modExcelSkills - Excel Manipulation Skills for z.ai Agent
' Defines tools and executes agent's tool calls on the active workbook
'==============================================================================
Option Explicit

' --- Get all tool definitions as JSON string ---
Public Function GetToolDefinitions() As String
    Dim tools As String
    tools = "[" & _
        ToolDef_ReadCell() & "," & _
        ToolDef_WriteCell() & "," & _
        ToolDef_ReadRange() & "," & _
        ToolDef_WriteRange() & "," & _
        ToolDef_GetSheetInfo() & "," & _
        ToolDef_GetWorkbookInfo() & "," & _
        ToolDef_FormatRange() & "," & _
        ToolDef_InsertFormula() & "," & _
        ToolDef_SortRange() & "," & _
        ToolDef_AddSheet() & "," & _
        ToolDef_DeleteRows() & "," & _
        ToolDef_InsertRows() & "," & _
        ToolDef_CreateChart() & _
    "]"
    GetToolDefinitions = tools
End Function

' --- Execute a tool call by name ---
Public Function ExecuteToolCall(ByVal toolName As String, ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim args As Object
    Set args = JsonParse(argsJson)
    If args Is Nothing Then Set args = CreateObject("Scripting.Dictionary")
    
    Dim result As String
    
    Select Case toolName
        Case "read_cell": result = SkillReadCell(args)
        Case "write_cell": result = SkillWriteCell(args)
        Case "read_range": result = SkillReadRange(args)
        Case "write_range": result = SkillWriteRange(args)
        Case "get_sheet_info": result = SkillGetSheetInfo(args)
        Case "get_workbook_info": result = SkillGetWorkbookInfo(args)
        Case "format_range": result = SkillFormatRange(args)
        Case "insert_formula": result = SkillInsertFormula(args)
        Case "sort_range": result = SkillSortRange(args)
        Case "add_sheet": result = SkillAddSheet(args)
        Case "delete_rows": result = SkillDeleteRows(args)
        Case "insert_rows": result = SkillInsertRows(args)
        Case "create_chart": result = SkillCreateChart(args)
        Case Else
            result = "{""error"":""Unknown tool: " & toolName & """}"
            LogWarn "Unknown tool call: " & toolName
    End Select
    
    LogToolCall toolName, argsJson, result
    ExecuteToolCall = result
    Exit Function
    
ErrHandler:
    Dim errMsg As String
    errMsg = "{""error"":""" & EscapeJsonString(Err.Description) & """}"
    LogErrorDetails "ExecuteToolCall:" & toolName, Err.Number, Err.Description
    ExecuteToolCall = errMsg
End Function

' ======================== SKILL IMPLEMENTATIONS ========================

Private Function GetTargetSheet(ByVal args As Object) As Worksheet
    If args.Exists("sheet") Then
        Dim shName As String
        shName = CStr(args("sheet"))
        If shName <> "" Then
            Set GetTargetSheet = ActiveWorkbook.Worksheets(shName)
            Exit Function
        End If
    End If
    Set GetTargetSheet = ActiveSheet
End Function

' --- Read Cell ---
Private Function SkillReadCell(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim cell As String
    cell = CStr(args("cell"))
    
    Dim rng As Range
    Set rng = ws.Range(cell)
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("cell") = cell
    result("value") = CStr(rng.value)
    result("formula") = rng.Formula
    result("type") = TypeName(rng.value)
    result("sheet") = ws.Name
    
    SkillReadCell = JsonStringify(result)
End Function

' --- Write Cell ---
Private Function SkillWriteCell(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim cell As String
    cell = CStr(args("cell"))
    
    Dim value As Variant
    value = args("value")
    
    ws.Range(cell).value = value
    
    SkillWriteCell = "{""success"":true,""cell"":""" & cell & """,""value"":""" & EscapeJsonString(CStr(value)) & """}"
End Function

' --- Read Range ---
Private Function SkillReadRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim rangeAddr As String
    rangeAddr = CStr(args("range"))
    
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    
    Dim result As String
    result = "{""range"":""" & rangeAddr & """,""sheet"":""" & ws.Name & """,""rows"":" & rng.Rows.Count & ",""cols"":" & rng.Columns.Count & ",""data"":["
    
    Dim r As Long, c As Long
    For r = 1 To rng.Rows.Count
        If r > 1 Then result = result & ","
        result = result & "["
        For c = 1 To rng.Columns.Count
            If c > 1 Then result = result & ","
            Dim cellVal As Variant
            cellVal = rng.Cells(r, c).value
            If IsEmpty(cellVal) Or IsNull(cellVal) Then
                result = result & "null"
            ElseIf IsNumeric(cellVal) And Not VarType(cellVal) = vbString Then
                result = result & Replace(CStr(cellVal), ",", ".")
            Else
                result = result & """" & EscapeJsonString(CStr(cellVal)) & """"
            End If
        Next c
        result = result & "]"
    Next r
    
    result = result & "]}"
    SkillReadRange = result
End Function

' --- Write Range ---
Private Function SkillWriteRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim startCell As String
    startCell = CStr(args("start_cell"))
    
    Dim data As Object
    Set data = args("data")
    
    Dim r As Long, c As Long
    Dim startRange As Range
    Set startRange = ws.Range(startCell)
    
    For r = 1 To data.Count
        Dim rowData As Object
        Set rowData = data(r)
        For c = 1 To rowData.Count
            startRange.Offset(r - 1, c - 1).value = rowData(c)
        Next c
    Next r
    
    SkillWriteRange = "{""success"":true,""start_cell"":""" & startCell & """,""rows_written"":" & data.Count & "}"
End Function

' --- Get Sheet Info ---
Private Function SkillGetSheetInfo(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim usedRng As Range
    Set usedRng = ws.UsedRange
    
    ' Read headers (first row)
    Dim headers As String
    headers = "["
    Dim c As Long
    Dim maxCols As Long
    maxCols = usedRng.Columns.Count
    If maxCols > 26 Then maxCols = 26
    For c = 1 To maxCols
        If c > 1 Then headers = headers & ","
        headers = headers & """" & EscapeJsonString(CStr(usedRng.Cells(1, c).value)) & """"
    Next c
    headers = headers & "]"
    
    ' Build JSON manually to avoid double-quoting the array
    Dim result As String
    result = "{" & _
        """name"":""" & EscapeJsonString(ws.Name) & """," & _
        """index"":" & ws.Index & "," & _
        """used_range"":""" & usedRng.Address & """," & _
        """used_rows"":" & usedRng.Rows.Count & "," & _
        """used_cols"":" & usedRng.Columns.Count & "," & _
        """first_cell"":""" & usedRng.Cells(1, 1).Address & """," & _
        """last_cell"":""" & usedRng.Cells(usedRng.Rows.Count, usedRng.Columns.Count).Address & """," & _
        """headers"":" & headers & _
    "}"
    
    SkillGetSheetInfo = result
End Function

' --- Get Workbook Info ---
Private Function SkillGetWorkbookInfo(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim result As String
    result = "{""name"":""" & EscapeJsonString(wb.Name) & """,""path"":""" & EscapeJsonString(wb.FullName) & """,""sheets"":["
    
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If i > 1 Then result = result & ","
        result = result & """" & EscapeJsonString(wb.Worksheets(i).Name) & """"
    Next i
    
    result = result & "],""active_sheet"":""" & EscapeJsonString(ActiveSheet.Name) & """}"
    SkillGetWorkbookInfo = result
End Function

' --- Format Range ---
Private Function SkillFormatRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim rangeAddr As String
    rangeAddr = CStr(args("range"))
    
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    
    If args.Exists("bold") Then rng.Font.Bold = CBool(args("bold"))
    If args.Exists("italic") Then rng.Font.Italic = CBool(args("italic"))
    If args.Exists("font_size") Then rng.Font.Size = CLng(args("font_size"))
    If args.Exists("font_color") Then rng.Font.Color = CLng(args("font_color"))
    If args.Exists("bg_color") Then rng.Interior.Color = CLng(args("bg_color"))
    If args.Exists("number_format") Then rng.NumberFormat = CStr(args("number_format"))
    If args.Exists("h_align") Then
        Select Case LCase(CStr(args("h_align")))
            Case "left": rng.HorizontalAlignment = xlLeft
            Case "center": rng.HorizontalAlignment = xlCenter
            Case "right": rng.HorizontalAlignment = xlRight
        End Select
    End If
    If args.Exists("wrap_text") Then rng.WrapText = CBool(args("wrap_text"))
    If args.Exists("borders") Then
        If CBool(args("borders")) Then
            rng.Borders.LineStyle = xlContinuous
            rng.Borders.Weight = xlThin
        Else
            rng.Borders.LineStyle = xlNone
        End If
    End If
    If args.Exists("column_width") Then rng.ColumnWidth = CDbl(args("column_width"))
    If args.Exists("row_height") Then rng.RowHeight = CDbl(args("row_height"))
    If args.Exists("autofit") Then
        If CBool(args("autofit")) Then
            rng.Columns.AutoFit
        End If
    End If
    If args.Exists("merge") Then
        If CBool(args("merge")) Then
            rng.Merge
        Else
            rng.UnMerge
        End If
    End If
    
    SkillFormatRange = "{""success"":true,""range"":""" & rangeAddr & """}"
End Function

' --- Insert Formula ---
Private Function SkillInsertFormula(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim cell As String
    cell = CStr(args("cell"))
    
    Dim formula As String
    formula = CStr(args("formula"))
    
    ' Ensure formula starts with =
    If Left(formula, 1) <> "=" Then formula = "=" & formula
    
    ws.Range(cell).Formula = formula
    
    Dim resultVal As String
    resultVal = CStr(ws.Range(cell).value)
    
    SkillInsertFormula = "{""success"":true,""cell"":""" & cell & """,""formula"":""" & EscapeJsonString(formula) & """,""result"":""" & EscapeJsonString(resultVal) & """}"
End Function

' --- Sort Range ---
Private Function SkillSortRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim rangeAddr As String
    rangeAddr = CStr(args("range"))
    
    Dim sortCol As String
    sortCol = CStr(args("sort_column"))
    
    Dim sortOrder As XlSortOrder
    If args.Exists("order") Then
        If LCase(CStr(args("order"))) = "desc" Or LCase(CStr(args("order"))) = "descending" Then
            sortOrder = xlDescending
        Else
            sortOrder = xlAscending
        End If
    Else
        sortOrder = xlAscending
    End If
    
    Dim hasHeaders As XlYesNoGuess
    If args.Exists("has_headers") Then
        hasHeaders = IIf(CBool(args("has_headers")), xlYes, xlNo)
    Else
        hasHeaders = xlYes
    End If
    
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    
    Dim sortKey As Range
    Set sortKey = ws.Range(sortCol & "1") ' Use column letter
    
    rng.Sort Key1:=ws.Range(sortCol & rng.Row), Order1:=sortOrder, Header:=hasHeaders
    
    SkillSortRange = "{""success"":true,""range"":""" & rangeAddr & """,""sort_column"":""" & sortCol & """}"
End Function

' --- Add Sheet ---
Private Function SkillAddSheet(ByVal args As Object) As String
    Dim sheetName As String
    If args.Exists("name") Then
        sheetName = CStr(args("name"))
    Else
        sheetName = ""
    End If
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    
    If sheetName <> "" Then
        On Error Resume Next
        ws.Name = sheetName
        On Error GoTo 0
    End If
    
    SkillAddSheet = "{""success"":true,""name"":""" & EscapeJsonString(ws.Name) & """}"
End Function

' --- Delete Rows ---
Private Function SkillDeleteRows(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim startRow As Long
    startRow = CLng(args("start_row"))
    
    Dim rowCount As Long
    If args.Exists("count") Then
        rowCount = CLng(args("count"))
    Else
        rowCount = 1
    End If
    
    ws.Rows(startRow & ":" & (startRow + rowCount - 1)).Delete
    
    SkillDeleteRows = "{""success"":true,""deleted_from"":" & startRow & ",""count"":" & rowCount & "}"
End Function

' --- Insert Rows ---
Private Function SkillInsertRows(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim atRow As Long
    atRow = CLng(args("at_row"))
    
    Dim rowCount As Long
    If args.Exists("count") Then
        rowCount = CLng(args("count"))
    Else
        rowCount = 1
    End If
    
    Dim i As Long
    For i = 1 To rowCount
        ws.Rows(atRow).Insert Shift:=xlDown
    Next i
    
    SkillInsertRows = "{""success"":true,""at_row"":" & atRow & ",""count"":" & rowCount & "}"
End Function

' --- Create Chart ---
Private Function SkillCreateChart(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim dataRange As String
    dataRange = CStr(args("data_range"))
    
    Dim chartType As XlChartType
    If args.Exists("chart_type") Then
        Select Case LCase(CStr(args("chart_type")))
            Case "bar": chartType = xlBarClustered
            Case "line": chartType = xlLine
            Case "pie": chartType = xlPie
            Case "scatter", "xy": chartType = xlXYScatter
            Case "area": chartType = xlArea
            Case "column": chartType = xlColumnClustered
            Case Else: chartType = xlColumnClustered
        End Select
    Else
        chartType = xlColumnClustered
    End If
    
    Dim chartTitle As String
    If args.Exists("title") Then
        chartTitle = CStr(args("title"))
    Else
        chartTitle = ""
    End If
    
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range(dataRange).Left + ws.Range(dataRange).Width + 20, _
        Top:=ws.Range(dataRange).Top, _
        Width:=400, Height:=300)
    
    With chartObj.Chart
        .SetSourceData Source:=ws.Range(dataRange)
        .chartType = chartType
        If chartTitle <> "" Then
            .HasTitle = True
            .ChartTitle.Text = chartTitle
        End If
    End With
    
    SkillCreateChart = "{""success"":true,""chart_name"":""" & EscapeJsonString(chartObj.Name) & """,""type"":""" & DictGet(args, "chart_type", "column") & """}"
End Function

' ======================== TOOL DEFINITIONS (JSON) ========================

Private Function ToolDef_ReadCell() As String
    ToolDef_ReadCell = "{""type"":""function"",""function"":{" & _
        """name"":""read_cell""," & _
        """description"":""Read value, formula and type from a single cell""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """cell"":{""type"":""string"",""description"":""Cell address e.g. A1, B2, C10""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional, defaults to active sheet)""}" & _
        "},""required"":[""cell""]}}}"
End Function

Private Function ToolDef_WriteCell() As String
    ToolDef_WriteCell = "{""type"":""function"",""function"":{" & _
        """name"":""write_cell""," & _
        """description"":""Write a value to a single cell""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """cell"":{""type"":""string"",""description"":""Cell address e.g. A1""}," & _
            """value"":{""type"":""string"",""description"":""Value to write""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""cell"",""value""]}}}"
End Function

Private Function ToolDef_ReadRange() As String
    ToolDef_ReadRange = "{""type"":""function"",""function"":{" & _
        """name"":""read_range""," & _
        """description"":""Read all values from a range of cells. Returns 2D array of values.""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """range"":{""type"":""string"",""description"":""Range address e.g. A1:D10""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""range""]}}}"
End Function

Private Function ToolDef_WriteRange() As String
    ToolDef_WriteRange = "{""type"":""function"",""function"":{" & _
        """name"":""write_range""," & _
        """description"":""Write a 2D array of values starting from a cell""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """start_cell"":{""type"":""string"",""description"":""Top-left cell to start writing e.g. A1""}," & _
            """data"":{""type"":""array"",""description"":""2D array of values, e.g. [[1,2],[3,4]]"",""items"":{""type"":""array"",""items"":{""type"":""string""}}}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""start_cell"",""data""]}}}"
End Function

Private Function ToolDef_GetSheetInfo() As String
    ToolDef_GetSheetInfo = "{""type"":""function"",""function"":{" & _
        """name"":""get_sheet_info""," & _
        """description"":""Get information about a worksheet: name, used range, dimensions, headers (first row)""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional, defaults to active sheet)""}" & _
        "},""required"":[]}}}"
End Function

Private Function ToolDef_GetWorkbookInfo() As String
    ToolDef_GetWorkbookInfo = "{""type"":""function"",""function"":{" & _
        """name"":""get_workbook_info""," & _
        """description"":""Get workbook information: file name, path, list of all sheet names, active sheet""," & _
        """parameters"":{""type"":""object"",""properties"":{}," & _
        """required"":[]}}}"
End Function

Private Function ToolDef_FormatRange() As String
    ToolDef_FormatRange = "{""type"":""function"",""function"":{" & _
        """name"":""format_range""," & _
        """description"":""Format cells: bold, italic, font size/color, background color, number format, alignment, borders, column width, row height, autofit, merge""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """range"":{""type"":""string"",""description"":""Range to format e.g. A1:D1""}," & _
            """bold"":{""type"":""boolean"",""description"":""Set bold""}," & _
            """italic"":{""type"":""boolean"",""description"":""Set italic""}," & _
            """font_size"":{""type"":""number"",""description"":""Font size in points""}," & _
            """font_color"":{""type"":""number"",""description"":""Font color as RGB long (e.g. 255 for red, 65280 for green, 16711680 for blue)""}," & _
            """bg_color"":{""type"":""number"",""description"":""Background color as RGB long""}," & _
            """number_format"":{""type"":""string"",""description"":""Number format string e.g. #,##0.00 or 0% or yyyy-mm-dd""}," & _
            """h_align"":{""type"":""string"",""description"":""Horizontal alignment: left, center, right""}," & _
            """wrap_text"":{""type"":""boolean"",""description"":""Enable text wrapping""}," & _
            """borders"":{""type"":""boolean"",""description"":""Add/remove thin borders""}," & _
            """column_width"":{""type"":""number"",""description"":""Set column width""}," & _
            """row_height"":{""type"":""number"",""description"":""Set row height""}," & _
            """autofit"":{""type"":""boolean"",""description"":""Auto-fit column width""}," & _
            """merge"":{""type"":""boolean"",""description"":""Merge/unmerge cells""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""range""]}}}"
End Function

Private Function ToolDef_InsertFormula() As String
    ToolDef_InsertFormula = "{""type"":""function"",""function"":{" & _
        """name"":""insert_formula""," & _
        """description"":""Insert an Excel formula into a cell. Use English function names (SUM, AVERAGE, VLOOKUP, IF, COUNT, etc.)""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """cell"":{""type"":""string"",""description"":""Target cell e.g. B10""}," & _
            """formula"":{""type"":""string"",""description"":""Formula e.g. =SUM(B1:B9) or =IF(A1>0,A1*2,0)""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""cell"",""formula""]}}}"
End Function

Private Function ToolDef_SortRange() As String
    ToolDef_SortRange = "{""type"":""function"",""function"":{" & _
        """name"":""sort_range""," & _
        """description"":""Sort a range of cells by a specified column""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """range"":{""type"":""string"",""description"":""Range to sort e.g. A1:D20""}," & _
            """sort_column"":{""type"":""string"",""description"":""Column letter to sort by e.g. B""}," & _
            """order"":{""type"":""string"",""description"":""Sort order: asc or desc (default: asc)""}," & _
            """has_headers"":{""type"":""boolean"",""description"":""Whether first row contains headers (default: true)""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""range"",""sort_column""]}}}"
End Function

Private Function ToolDef_AddSheet() As String
    ToolDef_AddSheet = "{""type"":""function"",""function"":{" & _
        """name"":""add_sheet""," & _
        """description"":""Add a new worksheet to the workbook""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """name"":{""type"":""string"",""description"":""Name for the new sheet (optional)""}" & _
        "},""required"":[]}}}"
End Function

Private Function ToolDef_DeleteRows() As String
    ToolDef_DeleteRows = "{""type"":""function"",""function"":{" & _
        """name"":""delete_rows""," & _
        """description"":""Delete one or more rows from the worksheet""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """start_row"":{""type"":""number"",""description"":""First row number to delete""}," & _
            """count"":{""type"":""number"",""description"":""Number of rows to delete (default: 1)""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""start_row""]}}}"
End Function

Private Function ToolDef_InsertRows() As String
    ToolDef_InsertRows = "{""type"":""function"",""function"":{" & _
        """name"":""insert_rows""," & _
        """description"":""Insert blank rows at a specified position""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """at_row"":{""type"":""number"",""description"":""Row number where to insert""}," & _
            """count"":{""type"":""number"",""description"":""Number of rows to insert (default: 1)""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""at_row""]}}}"
End Function

Private Function ToolDef_CreateChart() As String
    ToolDef_CreateChart = "{""type"":""function"",""function"":{" & _
        """name"":""create_chart""," & _
        """description"":""Create a chart from data range""," & _
        """parameters"":{""type"":""object"",""properties"":{" & _
            """data_range"":{""type"":""string"",""description"":""Data range for the chart e.g. A1:B10""}," & _
            """chart_type"":{""type"":""string"",""description"":""Chart type: column, bar, line, pie, scatter, area (default: column)""}," & _
            """title"":{""type"":""string"",""description"":""Chart title (optional)""}," & _
            """sheet"":{""type"":""string"",""description"":""Sheet name (optional)""}" & _
        "},""required"":[""data_range""]}}}"
End Function
