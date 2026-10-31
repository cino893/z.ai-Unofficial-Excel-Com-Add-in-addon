Attribute VB_Name = "modExcelSkillsOps"
'==============================================================================
' modExcelSkillsOps - Data, sheet, format, filter, chart, validation skills
'==============================================================================
Option Explicit

Private Const MAX_READ_ROWS As Long = 500
Private Const XL_TO_RIGHT As Long = -4161
Private Const XL_DOWN As Long = -4121
Private Const XL_PASTE_VALUES As Long = -4163
Private Const XL_UNDERLINE_SINGLE As Long = 2
Private Const XL_UNDERLINE_NONE As Long = -4142
Private Const XL_NONE As Long = -4142
Private Const XL_LEFT As Long = -4131
Private Const XL_CENTER As Long = -4108
Private Const XL_RIGHT As Long = -4152
Private Const XL_FILL_DEFAULT As Long = 0

' --- read_cell ---
Public Function SkillReadCell(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim cell As String
    cell = ArgStr(args, "cell")
    Dim rng As Range
    Set rng = ws.Range(cell)
    Dim valStr As String
    If IsEmpty(rng.value) Or IsNull(rng.value) Then
        valStr = ""
    Else
        valStr = CStr(rng.value)
    End If
    SkillReadCell = "{" & JsonStr("cell", cell) & "," & _
        JsonStr("value", valStr) & "," & _
        JsonStr("formula", CStr(rng.Formula)) & "," & _
        JsonStr("type", TypeName(rng.value)) & "," & _
        JsonStr("sheet", ws.Name) & "}"
End Function

' --- write_cell ---
Public Function SkillWriteCell(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim cell As String
    cell = ArgStr(args, "cell")
    Dim raw As String
    raw = ArgStr(args, "value")
    Dim num As Double
    If TryCoerceNumericCell(raw, num) Then
        ws.Range(cell).value = num
        SkillWriteCell = "{""success"":true," & JsonStr("cell", cell) & _
            ",""value"":" & Replace(CStr(num), ",", ".") & "}"
    Else
        ws.Range(cell).value = raw
        SkillWriteCell = "{""success"":true," & JsonStr("cell", cell) & "," & _
            JsonStr("value", raw) & "}"
    End If
End Function

' --- read_range ---
Public Function SkillReadRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    
    Dim totalRows As Long, cols As Long, rows As Long
    Dim truncated As Boolean
    totalRows = rng.Rows.Count
    cols = rng.Columns.Count
    truncated = (totalRows > MAX_READ_ROWS)
    If truncated Then rows = MAX_READ_ROWS Else rows = totalRows
    
    Dim dataJson As String
    dataJson = "["
    
    If rows = 1 And cols = 1 Then
        dataJson = dataJson & "[" & CellToJson(rng.value) & "]"
    Else
        Dim readRange As Range
        If truncated Then
            Set readRange = ws.Range(rng.Cells(1, 1), rng.Cells(rows, cols))
        Else
            Set readRange = rng
        End If
        Dim values As Variant
        values = readRange.value
        Dim r As Long, c As Long
        ' Multi-cell Range.Value is 2D (even for one row / one column); fall back to 1D.
        For r = 1 To rows
            If r > 1 Then dataJson = dataJson & ","
            dataJson = dataJson & "["
            For c = 1 To cols
                If c > 1 Then dataJson = dataJson & ","
                dataJson = dataJson & CellToJson(RangeValueAt(values, r, c))
            Next c
            dataJson = dataJson & "]"
        Next r
    End If
    dataJson = dataJson & "]"
    
    Dim result As String
    result = "{" & JsonStr("range", rangeAddr) & "," & JsonStr("sheet", ws.Name) & "," & _
        JsonNum("rows", rows) & "," & JsonNum("cols", cols) & ",""data"":" & dataJson
    If truncated Then
        result = result & ",""truncated"":true," & JsonNum("total_rows", totalRows)
    End If
    result = result & "}"
    SkillReadRange = result
End Function

Private Function RangeValueAt(ByVal values As Variant, ByVal r As Long, ByVal c As Long) As Variant
    On Error Resume Next
    RangeValueAt = values(r, c)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    If r = 1 Then
        RangeValueAt = values(c)
        If Err.Number = 0 Then Exit Function
        Err.Clear
    End If
    If c = 1 Then
        RangeValueAt = values(r)
        If Err.Number = 0 Then Exit Function
        Err.Clear
    End If
    RangeValueAt = Empty
End Function

Private Function CellToJson(ByVal cellVal As Variant) As String
    If IsEmpty(cellVal) Or IsNull(cellVal) Then
        CellToJson = "null"
    ElseIf IsNumeric(cellVal) And Not VarType(cellVal) = vbString Then
        CellToJson = Replace(CStr(cellVal), ",", ".")
    Else
        CellToJson = """" & EscapeJsonString(CStr(cellVal)) & """"
    End If
End Function

' --- write_range ---
Public Function SkillWriteRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim startCell As String
    startCell = ArgStr(args, "start_cell")
    Dim data As Collection
    Set data = ArgColl(args, "data")
    If data Is Nothing Then
        SkillWriteRange = "{""error"":""Missing 'data' parameter""}"
        Exit Function
    End If
    Dim rows As Long
    rows = data.Count
    If rows = 0 Then
        SkillWriteRange = "{""success"":true," & JsonStr("start_cell", startCell) & ",""rows_written"":0}"
        Exit Function
    End If
    Dim firstRow As Object
    Set firstRow = data(1)
    Dim cols As Long
    cols = firstRow.Count
    If cols = 0 Then
        SkillWriteRange = "{""success"":true," & JsonStr("start_cell", startCell) & ",""rows_written"":0}"
        Exit Function
    End If
    
    Dim values() As Variant
    ReDim values(1 To rows, 1 To cols)
    Dim r As Long, c As Long
    For r = 1 To rows
        Dim rowData As Object
        Set rowData = data(r)
        For c = 1 To cols
            If c <= rowData.Count Then
                Dim node As Variant
                node = rowData(c)
                If IsNull(node) Or IsEmpty(node) Then
                    values(r, c) = ""
                ElseIf IsNumeric(node) And Not VarType(node) = vbString Then
                    values(r, c) = node
                ElseIf VarType(node) = vbBoolean Then
                    values(r, c) = node
                Else
                    ' Keep JSON strings as text (parity with COM write_range).
                    values(r, c) = CStr(node)
                End If
            End If
        Next c
    Next r
    
    Dim startRange As Range
    Set startRange = ws.Range(startCell)
    Dim destRange As Range
    Set destRange = ws.Range(startRange, startRange.Offset(rows - 1, cols - 1))
    destRange.value = values
    
    SkillWriteRange = "{""success"":true," & JsonStr("start_cell", startCell) & "," & _
        JsonNum("rows_written", rows) & "," & JsonNum("cells_written", rows * cols) & "}"
End Function

' --- get_sheet_info ---
Public Function SkillGetSheetInfo(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim usedRng As Range
    Set usedRng = ws.UsedRange
    
    Dim usedRows As Long, usedCols As Long, maxCols As Long
    usedRows = usedRng.Rows.Count
    usedCols = usedRng.Columns.Count
    maxCols = usedCols
    If maxCols > 26 Then maxCols = 26
    
    Dim headers As String
    headers = "["
    Dim c As Long
    For c = 1 To maxCols
        If c > 1 Then headers = headers & ","
        headers = headers & """" & EscapeJsonString(CStr(usedRng.Cells(1, c).value)) & """"
    Next c
    headers = headers & "]"
    
    Dim sampleRows As Long
    sampleRows = usedRows
    If sampleRows > 5 Then sampleRows = 5
    Dim sample As String
    sample = "["
    Dim r As Long
    If sampleRows > 0 And maxCols > 0 Then
        For r = 1 To sampleRows
            If r > 1 Then sample = sample & ","
            sample = sample & "["
            For c = 1 To maxCols
                If c > 1 Then sample = sample & ","
                sample = sample & """" & EscapeJsonString(CStr(usedRng.Cells(r, c).value)) & """"
            Next c
            sample = sample & "]"
        Next r
    End If
    sample = sample & "]"
    
    Dim hasCharts As Boolean, hasFilters As Boolean
    Dim pivotsJson As String
    hasCharts = False
    hasFilters = False
    pivotsJson = "false"
    On Error Resume Next
    hasCharts = (ws.ChartObjects.Count > 0)
    Err.Clear
    Dim pCount As Long
    pCount = ws.PivotTables.Count
    If Err.Number <> 0 Then
        pivotsJson = """unknown"""
        Err.Clear
    Else
        If pCount > 0 Then pivotsJson = "true" Else pivotsJson = "false"
    End If
    hasFilters = ws.AutoFilterMode
    On Error GoTo 0
    
    SkillGetSheetInfo = "{" & _
        JsonStr("name", ws.Name) & "," & _
        JsonNum("index", ws.Index) & "," & _
        JsonStr("used_range", usedRng.Address) & "," & _
        JsonNum("used_rows", usedRows) & "," & _
        JsonNum("used_cols", usedCols) & "," & _
        JsonStr("first_cell", usedRng.Cells(1, 1).Address) & "," & _
        JsonStr("last_cell", usedRng.Cells(usedRows, usedCols).Address) & "," & _
        """headers"":" & headers & "," & _
        """sample_data"":" & sample & "," & _
        JsonBool("has_charts", hasCharts) & "," & _
        """has_pivot_tables"":" & pivotsJson & "," & _
        JsonBool("has_filters", hasFilters) & "}"
End Function

' --- get_workbook_info ---
Public Function SkillGetWorkbookInfo(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim sheetsJson As String
    sheetsJson = "["
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If i > 1 Then sheetsJson = sheetsJson & ","
        sheetsJson = sheetsJson & """" & EscapeJsonString(wb.Worksheets(i).Name) & """"
    Next i
    sheetsJson = sheetsJson & "]"
    
    Dim result As String
    result = "{" & JsonStr("name", wb.Name) & "," & JsonStr("path", wb.FullName) & _
        ",""sheets"":" & sheetsJson & "," & JsonStr("active_sheet", ActiveSheet.Name)
    
    On Error Resume Next
    Dim cacheCount As Long
    cacheCount = wb.PivotCaches.Count
    If Err.Number = 0 Then
        result = result & "," & JsonNum("pivot_cache_count", cacheCount)
    End If
    On Error GoTo 0
    
    result = result & "}"
    SkillGetWorkbookInfo = result
End Function

' --- format_range ---
Public Function SkillFormatRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    
    If ArgHas(args, "bold") Then rng.Font.Bold = ArgBool(args, "bold")
    If ArgHas(args, "italic") Then rng.Font.Italic = ArgBool(args, "italic")
    If ArgHas(args, "underline") Then
        If ArgBool(args, "underline") Then
            rng.Font.Underline = XL_UNDERLINE_SINGLE
        Else
            rng.Font.Underline = XL_UNDERLINE_NONE
        End If
    End If
    If ArgHas(args, "font_name") Then rng.Font.Name = ArgStr(args, "font_name")
    If ArgHas(args, "font_size") Then rng.Font.Size = ArgDbl(args, "font_size")
    If ArgHas(args, "font_color") Then rng.Font.Color = ArgLng(args, "font_color")
    If ArgHas(args, "bg_color") Then rng.Interior.Color = ArgLng(args, "bg_color")
    If ArgHas(args, "number_format") Then rng.NumberFormat = ArgStr(args, "number_format")
    If ArgHas(args, "h_align") Then
        Select Case LCase$(ArgStr(args, "h_align"))
            Case "center": rng.HorizontalAlignment = XL_CENTER
            Case "right": rng.HorizontalAlignment = XL_RIGHT
            Case Else: rng.HorizontalAlignment = XL_LEFT
        End Select
    End If
    If ArgHas(args, "wrap_text") Then rng.WrapText = ArgBool(args, "wrap_text")
    If ArgHas(args, "borders") Then
        If ArgBool(args, "borders") Then
            rng.Borders.LineStyle = xlContinuous
            rng.Borders.Weight = xlThin
        Else
            rng.Borders.LineStyle = XL_NONE
        End If
    End If
    If ArgHas(args, "column_width") Then rng.ColumnWidth = ArgDbl(args, "column_width")
    If ArgHas(args, "row_height") Then rng.RowHeight = ArgDbl(args, "row_height")
    If ArgHas(args, "autofit") Then
        If ArgBool(args, "autofit") Then rng.Columns.AutoFit
    End If
    If ArgHas(args, "merge") Then
        If ArgBool(args, "merge") Then rng.Merge Else rng.UnMerge
    End If
    
    SkillFormatRange = "{""success"":true," & JsonStr("range", rangeAddr) & "}"
End Function

' --- insert_formula ---
Public Function SkillInsertFormula(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim cell As String
    cell = ArgStr(args, "cell")
    Dim formula As String
    formula = ArgStr(args, "formula")
    Dim fillTo As String
    fillTo = ArgStr(args, "fill_to")
    
    If Left$(formula, 1) <> "=" Then formula = "=" & formula
    ws.Range(cell).Formula = formula
    
    Dim cellsFilled As Long
    cellsFilled = 1
    If fillTo <> "" Then
        Dim sourceRange As Range
        Set sourceRange = ws.Range(cell)
        Dim fillRange As Range
        Set fillRange = ws.Range(cell & ":" & fillTo)
        sourceRange.AutoFill Destination:=fillRange, Type:=XL_FILL_DEFAULT
        cellsFilled = fillRange.Cells.Count
    End If
    
    Dim resultVal As String
    If IsEmpty(ws.Range(cell).value) Then
        resultVal = ""
    Else
        resultVal = CStr(ws.Range(cell).value)
    End If
    
    SkillInsertFormula = "{""success"":true," & JsonStr("cell", cell) & "," & _
        JsonStr("formula", formula) & "," & JsonStr("result", resultVal) & "," & _
        JsonNum("cells_filled", cellsFilled) & "}"
End Function

' --- sort_range ---
Public Function SkillSortRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim sortCol As String
    sortCol = ArgStr(args, "sort_column")
    Dim orderStr As String
    orderStr = LCase$(ArgStr(args, "order", "asc"))
    Dim sortOrder As Long
    If orderStr = "desc" Or orderStr = "descending" Then sortOrder = xlDescending Else sortOrder = xlAscending
    Dim header As Long
    If ArgBool(args, "has_headers", True) Then header = xlYes Else header = xlNo
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    rng.Sort Key1:=ws.Range(sortCol & rng.Row), Order1:=sortOrder, Header:=header
    SkillSortRange = "{""success"":true," & JsonStr("range", rangeAddr) & "," & JsonStr("sort_column", sortCol) & "}"
End Function

' --- add_sheet ---
Public Function SkillAddSheet(ByVal args As Object) As String
    Dim requested As String
    requested = ArgStr(args, "name")
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    Dim actualName As String
    actualName = ws.Name
    Dim renamed As Boolean
    renamed = False
    If requested <> "" Then
        On Error Resume Next
        ws.Name = requested
        If Err.Number = 0 Then
            actualName = requested
            renamed = True
        End If
        On Error GoTo 0
    End If
    Dim result As String
    result = "{""success"":true," & JsonStr("name", actualName)
    If requested <> "" And Not renamed Then
        result = result & "," & JsonStr("warning", _
            "Could not rename sheet to '" & requested & "' (duplicate or invalid name). Created as '" & actualName & "'.")
    End If
    result = result & "}"
    SkillAddSheet = result
End Function

' --- delete_rows ---
Public Function SkillDeleteRows(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim startRow As Long
    startRow = ArgLng(args, "start_row")
    Dim rowCount As Long
    rowCount = ArgLng(args, "count", 1)
    If startRow < 1 Then
        SkillDeleteRows = "{""error"":""start_row must be >= 1""}"
        Exit Function
    End If
    If rowCount < 1 Then rowCount = 1
    ws.Rows(startRow & ":" & (startRow + rowCount - 1)).Delete
    SkillDeleteRows = "{""success"":true," & JsonNum("deleted_from", startRow) & "," & JsonNum("count", rowCount) & "}"
End Function

' --- insert_rows ---
Public Function SkillInsertRows(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim atRow As Long
    atRow = ArgLng(args, "at_row")
    Dim rowCount As Long
    rowCount = ArgLng(args, "count", 1)
    If atRow < 1 Then
        SkillInsertRows = "{""error"":""at_row must be >= 1""}"
        Exit Function
    End If
    If rowCount < 1 Then rowCount = 1
    ws.Rows(atRow & ":" & (atRow + rowCount - 1)).Insert Shift:=XL_DOWN
    SkillInsertRows = "{""success"":true," & JsonNum("at_row", atRow) & "," & JsonNum("count", rowCount) & "}"
End Function

' --- delete_columns ---
Public Function SkillDeleteColumns(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim startCol As String
    startCol = UCase$(ArgStr(args, "start_column"))
    Dim count As Long
    count = ArgLng(args, "count", 1)
    If startCol = "" Then
        SkillDeleteColumns = "{""error"":""start_column is required (e.g. 'C')""}"
        Exit Function
    End If
    If count < 1 Then count = 1
    Dim startColNum As Long, endColNum As Long
    startColNum = ColLetterToNum(startCol)
    endColNum = startColNum + count - 1
    Dim endCol As String
    endCol = ColNumToLetter(endColNum)
    ws.Columns(startCol & ":" & endCol).Delete
    SkillDeleteColumns = "{""success"":true," & JsonStr("deleted_from", startCol) & "," & _
        JsonStr("deleted_to", endCol) & "," & JsonNum("count", count) & "}"
End Function

' --- insert_columns ---
Public Function SkillInsertColumns(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim startCol As String
    startCol = UCase$(ArgStr(args, "start_column"))
    Dim count As Long
    count = ArgLng(args, "count", 1)
    If startCol = "" Then
        SkillInsertColumns = "{""error"":""start_column is required (e.g. 'C')""}"
        Exit Function
    End If
    If count < 1 Then count = 1
    Dim startColNum As Long, endColNum As Long
    startColNum = ColLetterToNum(startCol)
    endColNum = startColNum + count - 1
    Dim endCol As String
    endCol = ColNumToLetter(endColNum)
    ws.Columns(startCol & ":" & endCol).Insert Shift:=XL_TO_RIGHT
    SkillInsertColumns = "{""success"":true," & JsonStr("inserted_at", startCol) & "," & _
        JsonStr("inserted_to", endCol) & "," & JsonNum("count", count) & "}"
End Function

' --- create_chart ---
Public Function SkillCreateChart(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim dataRange As String
    dataRange = ArgStr(args, "data_range")
    Dim chartTypeStr As String
    chartTypeStr = LCase$(ArgStr(args, "chart_type", "column"))
    Dim title As String
    title = ArgStr(args, "title")
    
    Dim xlType As Long
    Select Case chartTypeStr
        Case "bar": xlType = xlBarClustered
        Case "line": xlType = xlLine
        Case "pie": xlType = xlPie
        Case "scatter", "xy": xlType = xlXYScatter
        Case "area": xlType = xlArea
        Case Else
            chartTypeStr = "column"
            xlType = xlColumnClustered
    End Select
    
    Dim srcRange As Range
    Set srcRange = ws.Range(dataRange)
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add( _
        Left:=srcRange.Left + srcRange.Width + 20, _
        Top:=srcRange.Top, Width:=400, Height:=300)
    chartObj.Chart.SetSourceData Source:=srcRange
    chartObj.Chart.chartType = xlType
    If title <> "" Then
        chartObj.Chart.HasTitle = True
        chartObj.Chart.ChartTitle.Text = title
    End If
    SkillCreateChart = "{""success"":true," & JsonStr("chart_name", chartObj.Name) & "," & JsonStr("type", chartTypeStr) & "}"
End Function

' --- delete_chart ---
Public Function SkillDeleteChart(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim chartName As String
    chartName = ArgStr(args, "chart_name")
    Dim i As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = chartName Then
            ws.ChartObjects(i).Delete
            SkillDeleteChart = "{""success"":true," & JsonStr("deleted", chartName) & "}"
            Exit Function
        End If
    Next i
    SkillDeleteChart = "{""error"":""Chart not found: " & EscapeJsonString(chartName) & """}"
End Function

' --- list_charts ---
Public Function SkillListCharts(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim count As Long
    count = ws.ChartObjects.Count
    Dim chartsJson As String
    chartsJson = "["
    Dim i As Long
    For i = 1 To count
        If i > 1 Then chartsJson = chartsJson & ","
        Dim co As ChartObject
        Set co = ws.ChartObjects(i)
        Dim typeName As String, source As String
        typeName = "unknown"
        source = ""
        On Error Resume Next
        Select Case co.Chart.chartType
            Case xlColumnClustered: typeName = "column"
            Case xlBarClustered: typeName = "bar"
            Case xlLine: typeName = "line"
            Case xlPie: typeName = "pie"
            Case xlXYScatter: typeName = "scatter"
            Case xlArea: typeName = "area"
            Case Else: typeName = "other(" & co.Chart.chartType & ")"
        End Select
        If co.Chart.SeriesCollection.Count > 0 Then
            source = CStr(co.Chart.SeriesCollection(1).Formula)
        End If
        On Error GoTo 0
        chartsJson = chartsJson & "{" & JsonStr("name", co.Name) & "," & _
            JsonStr("type", typeName) & "," & JsonStr("source", source) & "}"
    Next i
    chartsJson = chartsJson & "]"
    SkillListCharts = "{" & JsonStr("sheet", ws.Name) & "," & JsonNum("count", count) & ",""charts"":" & chartsJson & "}"
End Function

' --- auto_filter ---
Public Function SkillAutoFilter(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    
    If ArgBool(args, "clear") Then
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        SkillAutoFilter = "{""success"":true,""action"":""cleared""}"
        Exit Function
    End If
    
    Dim field As Long
    field = ArgLng(args, "field")
    Dim criteria As String
    criteria = ArgStr(args, "criteria")
    
    If field > 0 And criteria <> "" Then
        Dim op As String
        op = LCase$(ArgStr(args, "operator", "equals"))
        Select Case op
            Case "contains": rng.AutoFilter Field:=field, Criteria1:="*" & criteria & "*"
            Case "not_equals": rng.AutoFilter Field:=field, Criteria1:="<>" & criteria
            Case "greater": rng.AutoFilter Field:=field, Criteria1:=">" & criteria
            Case "less": rng.AutoFilter Field:=field, Criteria1:="<" & criteria
            Case "top10": rng.AutoFilter Field:=field, Criteria1:="10", Operator:=xlTop10Items
            Case "bottom10": rng.AutoFilter Field:=field, Criteria1:="10", Operator:=xlBottom10Items
            Case Else: rng.AutoFilter Field:=field, Criteria1:=criteria
        End Select
    Else
        rng.AutoFilter
    End If
    
    Dim action As String
    If field > 0 Then action = "filtered" Else action = "toggled"
    SkillAutoFilter = "{""success"":true," & JsonStr("range", rangeAddr) & "," & JsonStr("action", action) & "}"
End Function

' --- find_replace ---
Public Function SkillFindReplace(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim findStr As String, replaceStr As String, rangeAddr As String
    findStr = ArgStr(args, "find")
    replaceStr = ArgStr(args, "replace")
    rangeAddr = ArgStr(args, "range")
    Dim matchCase As Boolean, matchEntire As Boolean
    matchCase = ArgBool(args, "match_case")
    matchEntire = ArgBool(args, "match_entire")
    Dim searchRange As Range
    If rangeAddr = "" Then
        Set searchRange = ws.Cells
    Else
        Set searchRange = ws.Range(rangeAddr)
    End If
    Dim lookAt As Long
    If matchEntire Then lookAt = xlWhole Else lookAt = xlPart
    Dim replaced As Boolean
    replaced = searchRange.Replace(What:=findStr, Replacement:=replaceStr, LookAt:=lookAt, MatchCase:=matchCase)
    SkillFindReplace = "{""success"":" & LCase$(CStr(replaced)) & "," & _
        JsonStr("find", findStr) & "," & JsonStr("replace", replaceStr) & "}"
End Function

' --- conditional_format ---
Public Function SkillConditionalFormat(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    Dim ruleType As String
    ruleType = LCase$(ArgStr(args, "rule_type"))
    
    If ArgBool(args, "clear_existing") Then rng.FormatConditions.Delete
    
    Dim action As String
    action = ruleType
    Dim fc As Object
    
    Select Case ruleType
        Case "color_scale"
            rng.FormatConditions.AddColorScale ColorScaleType:=3
        Case "data_bars"
            rng.FormatConditions.AddDatabar
        Case "icon_set"
            rng.FormatConditions.AddIconSetCondition
        Case "duplicate"
            Set fc = rng.FormatConditions.AddUniqueValues
            fc.DupeUnique = xlDuplicate
            If ArgHas(args, "format_color") Then
                fc.Interior.Color = ArgLng(args, "format_color")
            Else
                fc.Interior.Color = 13551615
            End If
        Case "unique"
            Set fc = rng.FormatConditions.AddUniqueValues
            fc.DupeUnique = xlUnique
            If ArgHas(args, "format_color") Then
                fc.Interior.Color = ArgLng(args, "format_color")
            Else
                fc.Interior.Color = 5296274
            End If
        Case "above_average"
            Set fc = rng.FormatConditions.AddAboveAverage
            If ArgHas(args, "format_color") Then fc.Interior.Color = ArgLng(args, "format_color")
        Case "top_bottom"
            Dim tbVal As String, tbDir As String, tbRank As Long
            tbVal = ArgStr(args, "value1", "10")
            tbDir = LCase$(ArgStr(args, "operator", "top"))
            Set fc = rng.FormatConditions.AddTop10
            If tbDir = "bottom" Then fc.TopBottom = xlTop10Bottom Else fc.TopBottom = xlTop10Top
            If IsNumeric(tbVal) Then tbRank = CLng(Val(tbVal)) Else tbRank = 10
            If tbRank < 1 Then tbRank = 1
            If tbRank > 100 Then tbRank = 100
            fc.Rank = tbRank
            If ArgHas(args, "format_color") Then fc.Interior.Color = ArgLng(args, "format_color")
        Case "formula"
            Dim fml As String
            fml = ArgStr(args, "value1")
            If Left$(fml, 1) <> "=" Then fml = "=" & fml
            Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fml)
            If ArgHas(args, "format_color") Then fc.Interior.Color = ArgLng(args, "format_color")
            If ArgHas(args, "font_color") Then fc.Font.Color = ArgLng(args, "font_color")
        Case Else
            Dim opStr As String, val1 As String, val2 As String
            opStr = LCase$(ArgStr(args, "operator", "greater"))
            val1 = ArgStr(args, "value1", "0")
            val2 = ArgStr(args, "value2")
            Dim xlOp As Long
            Select Case opStr
                Case "less": xlOp = xlLess
                Case "equal": xlOp = xlEqual
                Case "between": xlOp = xlBetween
                Case "not_between": xlOp = xlNotBetween
                Case "greater_equal": xlOp = xlGreaterEqual
                Case "less_equal": xlOp = xlLessEqual
                Case Else: xlOp = xlGreater
            End Select
            If xlOp = xlBetween Or xlOp = xlNotBetween Then
                Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlOp, Formula1:=val1, Formula2:=val2)
            Else
                Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlOp, Formula1:=val1)
            End If
            If ArgHas(args, "format_color") Then fc.Interior.Color = ArgLng(args, "format_color")
            If ArgHas(args, "font_color") Then fc.Font.Color = ArgLng(args, "font_color")
            action = "cell_value"
    End Select
    
    SkillConditionalFormat = "{""success"":true," & JsonStr("range", rangeAddr) & "," & JsonStr("rule_type", action) & "}"
End Function

' --- copy_range ---
Public Function SkillCopyRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim sourceAddr As String, destAddr As String, destSheetName As String
    sourceAddr = ArgStr(args, "source")
    destAddr = ArgStr(args, "destination")
    destSheetName = ArgStr(args, "dest_sheet")
    Dim valuesOnly As Boolean
    valuesOnly = ArgBool(args, "values_only")
    Dim srcRange As Range
    Set srcRange = ws.Range(sourceAddr)
    Dim destSheet As Worksheet
    If destSheetName <> "" Then
        Set destSheet = ActiveWorkbook.Worksheets(destSheetName)
    Else
        Set destSheet = ws
    End If
    Dim destRange As Range
    Set destRange = destSheet.Range(destAddr)
    If valuesOnly Then
        srcRange.Copy
        destRange.PasteSpecial Paste:=XL_PASTE_VALUES
        Application.CutCopyMode = False
    Else
        srcRange.Copy destRange
    End If
    SkillCopyRange = "{""success"":true," & JsonStr("source", sourceAddr) & "," & JsonStr("destination", destAddr) & "}"
End Function

' --- rename_sheet ---
Public Function SkillRenameSheet(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim oldName As String, newName As String
    oldName = ws.Name
    newName = ArgStr(args, "new_name")
    ws.Name = newName
    SkillRenameSheet = "{""success"":true," & JsonStr("old_name", oldName) & "," & JsonStr("new_name", newName) & "}"
End Function

' --- delete_sheet ---
Public Function SkillDeleteSheet(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim sheetName As String
    sheetName = ArgStr(args, "sheet")
    If wb.Worksheets.Count <= 1 Then
        SkillDeleteSheet = "{""error"":""Cannot delete the last sheet in the workbook""}"
        Exit Function
    End If
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    Application.DisplayAlerts = False
    On Error GoTo RestoreAlerts
    ws.Delete
    Application.DisplayAlerts = True
    SkillDeleteSheet = "{""success"":true," & JsonStr("deleted", sheetName) & "}"
    Exit Function
RestoreAlerts:
    Application.DisplayAlerts = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' --- freeze_panes ---
Public Function SkillFreezePanes(ByVal args As Object) As String
    If ArgBool(args, "unfreeze") Then
        ActiveWindow.FreezePanes = False
        SkillFreezePanes = "{""success"":true,""action"":""unfrozen""}"
        Exit Function
    End If
    Dim cell As String, sheetName As String
    cell = ArgStr(args, "cell", "A2")
    sheetName = ArgStr(args, "sheet")
    If sheetName <> "" Then ActiveWorkbook.Worksheets(sheetName).Activate
    ActiveWindow.FreezePanes = False
    ActiveSheet.Range(cell).Select
    ActiveWindow.FreezePanes = True
    SkillFreezePanes = "{""success"":true,""action"":""frozen""," & JsonStr("cell", cell) & "}"
End Function

' --- remove_duplicates ---
Public Function SkillRemoveDuplicates(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    Dim hasHeaders As Boolean
    hasHeaders = ArgBool(args, "has_headers", True)
    Dim headerFlag As Long
    If hasHeaders Then headerFlag = xlYes Else headerFlag = xlNo
    
    Dim rowsBefore As Long, startRow As Long, col As Long
    rowsBefore = rng.Rows.Count
    startRow = rng.Row
    col = rng.Column
    Dim countRange As Range
    Set countRange = ws.Range(ws.Cells(startRow, col), ws.Cells(startRow + rowsBefore - 1, col))
    Dim dataBefore As Long, dataAfter As Long
    dataBefore = Application.WorksheetFunction.CountA(countRange)
    
    Dim colsNode As Collection
    Set colsNode = ArgColl(args, "columns")
    If Not colsNode Is Nothing Then
        If colsNode.Count > 0 Then
            Dim colArray() As Variant
            ReDim colArray(1 To colsNode.Count)
            Dim i As Long
            For i = 1 To colsNode.Count
                colArray(i) = CLng(colsNode(i))
            Next i
            rng.RemoveDuplicates Columns:=colArray, Header:=headerFlag
        Else
            ApplyRemoveDupAllCols rng, headerFlag
        End If
    Else
        ApplyRemoveDupAllCols rng, headerFlag
    End If
    
    dataAfter = Application.WorksheetFunction.CountA(countRange)
    SkillRemoveDuplicates = "{""success"":true," & JsonStr("range", rangeAddr) & "," & _
        JsonNum("rows_before", dataBefore) & "," & JsonNum("rows_after", dataAfter) & "," & _
        JsonNum("rows_removed", dataBefore - dataAfter) & "}"
End Function

Private Sub ApplyRemoveDupAllCols(ByVal rng As Range, ByVal headerFlag As Long)
    Dim colCount As Long, i As Long
    colCount = rng.Columns.Count
    Dim allCols() As Variant
    ReDim allCols(1 To colCount)
    For i = 1 To colCount
        allCols(i) = i
    Next i
    rng.RemoveDuplicates Columns:=allCols, Header:=headerFlag
End Sub

' --- set_validation ---
Public Function SkillSetValidation(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    Dim typeStr As String, formula1 As String, formula2 As String, opStr As String
    typeStr = LCase$(ArgStr(args, "type"))
    formula1 = ArgStr(args, "formula1")
    formula2 = ArgStr(args, "formula2")
    opStr = LCase$(ArgStr(args, "operator", "between"))
    
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    
    Dim xlType As Long
    Select Case typeStr
        Case "whole_number": xlType = xlValidateWholeNumber
        Case "decimal": xlType = xlValidateDecimal
        Case "date": xlType = xlValidateDate
        Case "text_length": xlType = xlValidateTextLength
        Case Else
            typeStr = "list"
            xlType = xlValidateList
    End Select
    
    Dim xlOp As Long
    Select Case opStr
        Case "not_between": xlOp = xlNotBetween
        Case "equal": xlOp = xlEqual
        Case "not_equal": xlOp = xlNotEqual
        Case "greater": xlOp = xlGreater
        Case "less": xlOp = xlLess
        Case "greater_equal": xlOp = xlGreaterEqual
        Case "less_equal": xlOp = xlLessEqual
        Case Else: xlOp = xlBetween
    End Select
    
    If xlType = xlValidateList Then
        rng.Validation.Add Type:=xlType, AlertStyle:=xlValidAlertStop, Formula1:=formula1
        rng.Validation.InCellDropdown = ArgBool(args, "show_dropdown", True)
    ElseIf formula2 <> "" Then
        rng.Validation.Add Type:=xlType, AlertStyle:=xlValidAlertStop, Operator:=xlOp, Formula1:=formula1, Formula2:=formula2
    Else
        rng.Validation.Add Type:=xlType, AlertStyle:=xlValidAlertStop, Operator:=xlOp, Formula1:=formula1
    End If
    
    Dim errorMsg As String
    errorMsg = ArgStr(args, "error_message")
    If errorMsg <> "" Then
        rng.Validation.ErrorMessage = errorMsg
        rng.Validation.ShowError = True
    End If
    
    SkillSetValidation = "{""success"":true," & JsonStr("range", rangeAddr) & "," & JsonStr("type", typeStr) & "}"
End Function

' --- clear_range ---
Public Function SkillClearRange(ByVal args As Object) As String
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim rangeAddr As String
    rangeAddr = ArgStr(args, "range")
    Dim what As String
    what = LCase$(ArgStr(args, "what", "contents"))
    Dim rng As Range
    Set rng = ws.Range(rangeAddr)
    Select Case what
        Case "formats": rng.ClearFormats
        Case "all": rng.Clear
        Case Else
            what = "contents"
            rng.ClearContents
    End Select
    SkillClearRange = "{""success"":true," & JsonStr("range", rangeAddr) & "," & JsonStr("cleared", what) & "}"
End Function
