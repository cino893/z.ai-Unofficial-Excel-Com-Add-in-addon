Attribute VB_Name = "modExcelSkills"
'==============================================================================
' modExcelSkills - Excel tool definitions, dispatch, and shared helpers
' Implementations live in modExcelSkillsOps / modExcelSkillsPivot
'==============================================================================
Option Explicit

Private Const VBA_E_IGNORE As Long = -2146777998   ' 0x800AC472 Excel busy / VBA_E_IGNORE
Private Const COM_E_INVALID As Long = -2146827284  ' 0x800A03EC (often surfaces as Err 1004)
Private Const TYPE_E_INVDATAREAD As Long = -2147319784 ' 0x80028018

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' --- Get all tool definitions as JSON string ---
Public Function GetToolDefinitions() As String
    Dim parts As New Collection
    parts.Add ToolDef_ReadCell()
    parts.Add ToolDef_WriteCell()
    parts.Add ToolDef_ReadRange()
    parts.Add ToolDef_WriteRange()
    parts.Add ToolDef_GetSheetInfo()
    parts.Add ToolDef_GetWorkbookInfo()
    parts.Add ToolDef_FormatRange()
    parts.Add ToolDef_InsertFormula()
    parts.Add ToolDef_SortRange()
    parts.Add ToolDef_AddSheet()
    parts.Add ToolDef_DeleteRows()
    parts.Add ToolDef_InsertRows()
    parts.Add ToolDef_DeleteColumns()
    parts.Add ToolDef_InsertColumns()
    parts.Add ToolDef_CreateChart()
    parts.Add ToolDef_DeleteChart()
    parts.Add ToolDef_ListCharts()
    parts.Add ToolDef_CreatePivotTable()
    parts.Add ToolDef_MovePivotTable()
    parts.Add ToolDef_ModifyPivotTable()
    parts.Add ToolDef_AutoFilter()
    parts.Add ToolDef_FindReplace()
    parts.Add ToolDef_ConditionalFormat()
    parts.Add ToolDef_CopyRange()
    parts.Add ToolDef_RenameSheet()
    parts.Add ToolDef_DeleteSheet()
    parts.Add ToolDef_FreezePanes()
    parts.Add ToolDef_RemoveDuplicates()
    parts.Add ToolDef_SetValidation()
    parts.Add ToolDef_ListPivotTables()
    parts.Add ToolDef_ClearRange()
    
    Dim i As Long
    Dim result As String
    result = "["
    For i = 1 To parts.Count
        If i > 1 Then result = result & ","
        result = result & parts(i)
    Next i
    result = result & "]"
    GetToolDefinitions = result
End Function

' --- Execute a tool call by name (retry once on Excel busy) ---
Public Function ExecuteToolCall(ByVal toolName As String, ByVal argsJson As String) As String
    Dim attempt As Long
    Dim result As String
    Dim errNum As Long
    Dim errDesc As String
    
    For attempt = 0 To 1
        On Error Resume Next
        Err.Clear
        result = ExecuteInner(toolName, argsJson)
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0
        
        If errNum = 0 Then
            LogToolCall toolName, argsJson, result
            ExecuteToolCall = result
            Exit Function
        End If
        
        If attempt = 0 And IsExcelBusyError(errNum, errDesc) Then
            LogWarn toolName & ": Excel busy (" & errNum & "), retrying in 300ms..."
            Sleep 300
        Else
            result = BuildToolErrorJson(errNum, errDesc)
            LogErrorDetails "ExecuteToolCall:" & toolName, errNum, errDesc
            LogToolCall toolName, argsJson, result
            ExecuteToolCall = result
            Exit Function
        End If
    Next attempt
    
    On Error Resume Next
    Err.Clear
    result = ExecuteInner(toolName, argsJson)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNum <> 0 Then
        result = BuildToolErrorJson(errNum, errDesc)
        LogErrorDetails "ExecuteToolCall:" & toolName, errNum, errDesc
    End If
    LogToolCall toolName, argsJson, result
    ExecuteToolCall = result
End Function

Private Function ExecuteInner(ByVal toolName As String, ByVal argsJson As String) As String
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
        Case "delete_columns": result = SkillDeleteColumns(args)
        Case "insert_columns": result = SkillInsertColumns(args)
        Case "create_chart": result = SkillCreateChart(args)
        Case "delete_chart": result = SkillDeleteChart(args)
        Case "list_charts": result = SkillListCharts(args)
        Case "create_pivot_table": result = SkillCreatePivotTable(args)
        Case "move_pivot_table": result = SkillMovePivotTable(args)
        Case "modify_pivot_table": result = SkillModifyPivotTable(args)
        Case "auto_filter": result = SkillAutoFilter(args)
        Case "find_replace": result = SkillFindReplace(args)
        Case "conditional_format": result = SkillConditionalFormat(args)
        Case "copy_range": result = SkillCopyRange(args)
        Case "rename_sheet": result = SkillRenameSheet(args)
        Case "delete_sheet": result = SkillDeleteSheet(args)
        Case "freeze_panes": result = SkillFreezePanes(args)
        Case "remove_duplicates": result = SkillRemoveDuplicates(args)
        Case "set_validation": result = SkillSetValidation(args)
        Case "list_pivot_tables": result = SkillListPivotTables(args)
        Case "clear_range": result = SkillClearRange(args)
        Case Else
            result = "{""error"":""Unknown tool: " & EscapeJsonString(toolName) & """}"
            LogWarn "Unknown tool call: " & toolName
    End Select
    ExecuteInner = result
End Function

Private Function IsExcelBusyError(ByVal errNum As Long, ByVal errDesc As String) As Boolean
    If errNum = VBA_E_IGNORE Then
        IsExcelBusyError = True
        Exit Function
    End If
    If errNum = 1004 Or errNum = COM_E_INVALID Then
        If InStr(1, errDesc, "busy", vbTextCompare) > 0 Then
            IsExcelBusyError = True
            Exit Function
        End If
    End If
    IsExcelBusyError = False
End Function

Private Function BuildToolErrorJson(ByVal errNum As Long, ByVal errDesc As String) As String
    Dim hint As String
    Dim descL As String
    descL = LCase$(errDesc)
    
    If InStr(descL, "pivot") > 0 Or InStr(descL, "tabela przestawna") > 0 Then
        hint = "A PivotTable is blocking this operation. Use move_pivot_table to relocate it to another sheet first, then retry."
    ElseIf InStr(descL, "merge") > 0 Or InStr(descL, "scal") > 0 Then
        hint = "Merged cells may be interfering. Try clear_range with what='formats' first, or work with a different range."
    ElseIf errNum = TYPE_E_INVDATAREAD Then
        hint = "TYPE_E_INVDATAREAD — Excel type library error. This sheet may have COM interop issues. Try a different approach or skip this sheet."
    ElseIf errNum = 1004 Or errNum = COM_E_INVALID Then
        hint = "COM error — the range may be invalid, protected, or overlap an existing object. Check the range address and try again."
    End If
    
    Dim json As String
    json = "{""error"":""" & EscapeJsonString(errDesc) & """"
    If hint <> "" Then json = json & ",""hint"":""" & EscapeJsonString(hint) & """"
    json = json & "}"
    BuildToolErrorJson = json
End Function

' ======================== SHARED HELPERS ========================

Public Function GetTargetSheet(ByVal args As Object) As Worksheet
    Dim shName As String
    shName = ArgStr(args, "sheet")
    If shName <> "" Then
        Set GetTargetSheet = ActiveWorkbook.Worksheets(shName)
        Exit Function
    End If
    Set GetTargetSheet = ActiveSheet
End Function

Public Function ColLetterToNum(ByVal col As String) As Long
    Dim i As Long, n As Long, ch As String
    col = UCase$(Trim$(col))
    n = 0
    For i = 1 To Len(col)
        ch = Mid$(col, i, 1)
        If ch >= "A" And ch <= "Z" Then
            n = n * 26 + (Asc(ch) - 64)
        End If
    Next i
    ColLetterToNum = n
End Function

Public Function ColNumToLetter(ByVal colNum As Long) As String
    Dim result As String
    Dim n As Long
    n = colNum
    result = ""
    Do While n > 0
        n = n - 1
        result = Chr$(65 + (n Mod 26)) & result
        n = n \ 26
    Loop
    ColNumToLetter = result
End Function

Public Function JsonEsc(ByVal s As String) As String
    JsonEsc = EscapeJsonString(s)
End Function

Public Function JsonStr(ByVal key As String, ByVal value As String) As String
    JsonStr = """" & key & """:""" & EscapeJsonString(value) & """"
End Function

Public Function JsonNum(ByVal key As String, ByVal value As Variant) As String
    JsonNum = """" & key & """:" & Replace(CStr(value), ",", ".")
End Function

Public Function JsonBool(ByVal key As String, ByVal value As Boolean) As String
    JsonBool = """" & key & """:" & IIf(value, "true", "false")
End Function

Public Function XlValueFunction(ByVal funcStr As String) As Long
    Select Case LCase$(Trim$(funcStr))
        Case "count": XlValueFunction = -4112
        Case "average": XlValueFunction = -4106
        Case "max": XlValueFunction = -4136
        Case "min": XlValueFunction = -4139
        Case Else: XlValueFunction = -4157 ' xlSum
    End Select
End Function

' ======================== TOOL DEFINITIONS (JSON) ========================

Private Function TD(ByVal name As String, ByVal desc As String, ByVal props As String, ByVal required As String) As String
    TD = "{""type"":""function"",""function"":{""name"":""" & name & """," & _
         """description"":""" & desc & """," & _
         """parameters"":{""type"":""object"",""properties"":{" & props & "}," & _
         """required"":[" & required & "]}}}"
End Function

Private Function PS(ByVal name As String, ByVal desc As String) As String
    PS = """" & name & """:{""type"":""string"",""description"":""" & desc & """}"
End Function

Private Function PN(ByVal name As String, ByVal desc As String) As String
    PN = """" & name & """:{""type"":""number"",""description"":""" & desc & """}"
End Function

Private Function PB(ByVal name As String, ByVal desc As String) As String
    PB = """" & name & """:{""type"":""boolean"",""description"":""" & desc & """}"
End Function

Private Function PA(ByVal name As String, ByVal desc As String, ByVal itemType As String) As String
    PA = """" & name & """:{""type"":""array"",""description"":""" & desc & """,""items"":{""type"":""" & itemType & """}}"
End Function

Private Function ToolDef_ReadCell() As String
    ToolDef_ReadCell = TD("read_cell", _
        "Read value, formula and type from a single cell", _
        PS("cell", "Cell address e.g. A1, B2, C10") & "," & _
        PS("sheet", "Sheet name (optional, defaults to active sheet)"), _
        """cell""")
End Function

Private Function ToolDef_WriteCell() As String
    ToolDef_WriteCell = TD("write_cell", _
        "Write a value to a single cell", _
        PS("cell", "Cell address e.g. A1") & "," & _
        PS("value", "Value to write") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """cell"",""value""")
End Function

Private Function ToolDef_ReadRange() As String
    ToolDef_ReadRange = TD("read_range", _
        "Read all values from a range of cells. Returns 2D array of values.", _
        PS("range", "Range address e.g. A1:D10") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range""")
End Function

Private Function ToolDef_WriteRange() As String
    ToolDef_WriteRange = TD("write_range", _
        "Write a 2D array of values starting from a cell. PREFERRED over multiple write_cell calls for bulk data.", _
        PS("start_cell", "Top-left cell to start writing e.g. A1") & "," & _
        """data"":{""type"":""array"",""description"":""2D array of values, e.g. [[1,2],[3,4]]"",""items"":{""type"":""array"",""items"":{""type"":""string""}}}," & _
        PS("sheet", "Sheet name (optional)"), _
        """start_cell"",""data""")
End Function

Private Function ToolDef_GetSheetInfo() As String
    ToolDef_GetSheetInfo = TD("get_sheet_info", _
        "Get worksheet info: name, used range, dimensions, headers, data sample (first 5 rows), and flags for charts/pivot tables/filters", _
        PS("sheet", "Sheet name (optional, defaults to active sheet)"), _
        "")
End Function

Private Function ToolDef_GetWorkbookInfo() As String
    ToolDef_GetWorkbookInfo = TD("get_workbook_info", _
        "Get workbook information: file name, path, list of all sheet names, active sheet", _
        "", "")
End Function

Private Function ToolDef_FormatRange() As String
    ToolDef_FormatRange = TD("format_range", _
        "Format cells: bold, italic, underline, font name/size/color, background color, number format, alignment, borders, column width, row height, autofit, merge", _
        PS("range", "Range to format e.g. A1:D1") & "," & _
        PB("bold", "Set bold") & "," & _
        PB("italic", "Set italic") & "," & _
        PB("underline", "Set underline") & "," & _
        PS("font_name", "Font name e.g. Arial, Calibri, Times New Roman") & "," & _
        PN("font_size", "Font size in points") & "," & _
        PN("font_color", "Font color as RGB long (e.g. 255 for red, 65280 for green, 16711680 for blue)") & "," & _
        PN("bg_color", "Background color as RGB long") & "," & _
        PS("number_format", "Number format string e.g. #,##0.00 or 0% or yyyy-mm-dd") & "," & _
        PS("h_align", "Horizontal alignment: left, center, right") & "," & _
        PB("wrap_text", "Enable text wrapping") & "," & _
        PB("borders", "Add/remove thin borders") & "," & _
        PN("column_width", "Set column width") & "," & _
        PN("row_height", "Set row height") & "," & _
        PB("autofit", "Auto-fit column width") & "," & _
        PB("merge", "Merge/unmerge cells") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range""")
End Function

Private Function ToolDef_InsertFormula() As String
    ToolDef_InsertFormula = TD("insert_formula", _
        "Insert an Excel formula into a cell. Use English function names (SUM, AVERAGE, VLOOKUP, IF, COUNT, etc.). Use fill_to to auto-fill the formula down/across a range.", _
        PS("cell", "Target cell e.g. B10") & "," & _
        PS("formula", "Formula e.g. =SUM(B1:B9) or =IF(A1>0,A1*2,0)") & "," & _
        PS("fill_to", "End cell to auto-fill formula to e.g. B100 (optional — fills from cell to fill_to adjusting references)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """cell"",""formula""")
End Function

Private Function ToolDef_SortRange() As String
    ToolDef_SortRange = TD("sort_range", _
        "Sort a range of cells by a specified column", _
        PS("range", "Range to sort e.g. A1:D20") & "," & _
        PS("sort_column", "Column letter to sort by e.g. B") & "," & _
        PS("order", "Sort order: asc or desc (default: asc)") & "," & _
        PB("has_headers", "Whether first row contains headers (default: true)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range"",""sort_column""")
End Function

Private Function ToolDef_AddSheet() As String
    ToolDef_AddSheet = TD("add_sheet", _
        "Add a new worksheet to the workbook", _
        PS("name", "Name for the new sheet (optional)"), _
        "")
End Function

Private Function ToolDef_DeleteRows() As String
    ToolDef_DeleteRows = TD("delete_rows", _
        "Delete one or more rows from the worksheet", _
        PN("start_row", "First row number to delete") & "," & _
        PN("count", "Number of rows to delete (default: 1)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """start_row""")
End Function

Private Function ToolDef_InsertRows() As String
    ToolDef_InsertRows = TD("insert_rows", _
        "Insert blank rows at a specified position", _
        PN("at_row", "Row number where to insert") & "," & _
        PN("count", "Number of rows to insert (default: 1)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """at_row""")
End Function

Private Function ToolDef_DeleteColumns() As String
    ToolDef_DeleteColumns = TD("delete_columns", _
        "Delete one or more columns from the worksheet", _
        PS("start_column", "First column letter to delete e.g. C") & "," & _
        PN("count", "Number of columns to delete (default: 1)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """start_column""")
End Function

Private Function ToolDef_InsertColumns() As String
    ToolDef_InsertColumns = TD("insert_columns", _
        "Insert one or more blank columns at a specified position", _
        PS("start_column", "First column letter where to insert e.g. C") & "," & _
        PN("count", "Number of columns to insert (default: 1)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """start_column""")
End Function

Private Function ToolDef_CreateChart() As String
    ToolDef_CreateChart = TD("create_chart", _
        "Create a chart from data range. Use list_charts first to check existing charts, and delete_chart to remove old ones before creating new.", _
        PS("data_range", "Data range for the chart e.g. A1:B10") & "," & _
        PS("chart_type", "Chart type: column, bar, line, pie, scatter, area (default: column)") & "," & _
        PS("title", "Chart title (optional)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """data_range""")
End Function

Private Function ToolDef_DeleteChart() As String
    ToolDef_DeleteChart = TD("delete_chart", _
        "Delete a chart from the worksheet by name. Use list_charts to find chart names.", _
        PS("chart_name", "Name of the chart to delete e.g. Chart 1") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """chart_name""")
End Function

Private Function ToolDef_ListCharts() As String
    ToolDef_ListCharts = TD("list_charts", _
        "List all charts on a worksheet with their names, types and data ranges", _
        PS("sheet", "Sheet name (optional, defaults to active sheet)"), _
        "")
End Function

Private Function ToolDef_CreatePivotTable() As String
    ToolDef_CreatePivotTable = TD("create_pivot_table", _
        "Create a PivotTable from a data range. Specify row, column and value fields.", _
        PS("source_range", "Source data range e.g. A1:D100") & "," & _
        PS("dest_cell", "Top-left cell for PivotTable placement e.g. F1 (default: new sheet)") & "," & _
        PS("dest_sheet", "Destination sheet name (optional, defaults to source sheet or new sheet)") & "," & _
        PS("name", "PivotTable name (optional)") & "," & _
        PA("row_fields", "Fields to use as row labels e.g. Category, Region", "string") & "," & _
        PA("column_fields", "Fields to use as column labels (optional)", "string") & "," & _
        PA("value_fields", "Fields to aggregate as values e.g. Sales, Quantity", "string") & "," & _
        PS("value_function", "Aggregation function: sum, count, average, max, min (default: sum)") & "," & _
        PS("sheet", "Source sheet name (optional)"), _
        """source_range"",""row_fields"",""value_fields""")
End Function

Private Function ToolDef_MovePivotTable() As String
    ToolDef_MovePivotTable = TD("move_pivot_table", _
        "Move a PivotTable to another sheet. Tries Location property, then recreates from PivotCache. On .xlsb files, recreated pivot is empty — use modify_pivot_table to configure fields afterward.", _
        PS("name", "PivotTable name to move (use list_pivot_tables to find it)") & "," & _
        PS("source_range", "Range containing the pivot e.g. H1:M30 (alternative to name, recommended for .xlsb)") & "," & _
        PS("dest_sheet", "Destination sheet name (creates if not exists)") & "," & _
        PS("dest_cell", "Destination cell e.g. A1 (default: A1)") & "," & _
        PS("sheet", "Source sheet name (optional)"), _
        "")
End Function

Private Function ToolDef_ModifyPivotTable() As String
    ToolDef_ModifyPivotTable = TD("modify_pivot_table", _
        "Configure fields on an existing PivotTable: set row/column/value/page fields, change aggregation function, refresh. Use after create_pivot_table or move_pivot_table to set up field layout.", _
        PS("name", "PivotTable name (optional, defaults to first pivot on sheet)") & "," & _
        PA("row_fields", "Fields for row labels e.g. Category, Region", "string") & "," & _
        PA("column_fields", "Fields for column labels (optional)", "string") & "," & _
        PA("value_fields", "Fields to aggregate e.g. Sales", "string") & "," & _
        PS("value_function", "Aggregation: sum, count, average, max, min (default: sum)") & "," & _
        PA("page_fields", "Fields for page/filter area (optional)", "string") & "," & _
        PB("clear_fields", "Reset all field orientations before applying new ones (default: true)") & "," & _
        PB("refresh", "Refresh the PivotTable after changes (default: false)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        "")
End Function

Private Function ToolDef_AutoFilter() As String
    ToolDef_AutoFilter = TD("auto_filter", _
        "Apply or remove auto-filter on a range. Call without criteria to toggle filter on/off.", _
        PS("range", "Range to filter e.g. A1:D20") & "," & _
        PN("field", "Column number within range to filter (1-based, optional)") & "," & _
        PS("criteria", "Filter criteria value (optional, omit to show all or toggle)") & "," & _
        PS("operator", "Filter operator: equals, not_equals, greater, less, contains, top10, bottom10 (default: equals)") & "," & _
        PB("clear", "Set true to clear all filters") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range""")
End Function

Private Function ToolDef_FindReplace() As String
    ToolDef_FindReplace = TD("find_replace", _
        "Find and replace values in a range or entire sheet", _
        PS("find", "Text to find") & "," & _
        PS("replace", "Replacement text") & "," & _
        PS("range", "Range to search (optional, defaults to entire sheet)") & "," & _
        PB("match_case", "Case-sensitive matching (default: false)") & "," & _
        PB("match_entire", "Match entire cell contents (default: false)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """find"",""replace""")
End Function

Private Function ToolDef_ConditionalFormat() As String
    ToolDef_ConditionalFormat = TD("conditional_format", _
        "Add conditional formatting to a range. Supports color scale, data bars, icon sets, formula-based rules, and value-based rules.", _
        PS("range", "Range to format e.g. B2:B100") & "," & _
        PS("rule_type", "Rule type: color_scale, data_bars, icon_set, cell_value, formula, top_bottom, above_average, duplicate, unique. Use 'formula' for alternating row colors (=MOD(ROW(),2)=0), row highlighting, date comparisons.") & "," & _
        PS("operator", "For cell_value: greater, less, equal, between, not_between, greater_equal, less_equal. For top_bottom: top or bottom (default: top)") & "," & _
        PS("value1", "Threshold value, or Excel formula for rule_type=formula (e.g. =MOD(ROW(),2)=0)") & "," & _
        PS("value2", "Second value (for between/not_between operator)") & "," & _
        PN("format_color", "Fill color as RGB long for matching cells (e.g. 65280 for green)") & "," & _
        PN("font_color", "Font color as RGB long for matching cells") & "," & _
        PB("clear_existing", "Clear existing conditional formats before adding (default: false)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range"",""rule_type""")
End Function

Private Function ToolDef_CopyRange() As String
    ToolDef_CopyRange = TD("copy_range", _
        "Copy a range of cells to another location (values, formulas and formatting)", _
        PS("source", "Source range e.g. A1:D10") & "," & _
        PS("destination", "Destination top-left cell e.g. F1") & "," & _
        PS("dest_sheet", "Destination sheet name (optional, defaults to same sheet)") & "," & _
        PB("values_only", "Paste values only, no formulas/formatting (default: false)") & "," & _
        PS("sheet", "Source sheet name (optional)"), _
        """source"",""destination""")
End Function

Private Function ToolDef_RenameSheet() As String
    ToolDef_RenameSheet = TD("rename_sheet", _
        "Rename a worksheet", _
        PS("sheet", "Current sheet name (optional, defaults to active sheet)") & "," & _
        PS("new_name", "New name for the sheet"), _
        """new_name""")
End Function

Private Function ToolDef_DeleteSheet() As String
    ToolDef_DeleteSheet = TD("delete_sheet", _
        "Delete a worksheet from the workbook", _
        PS("sheet", "Sheet name to delete"), _
        """sheet""")
End Function

Private Function ToolDef_FreezePanes() As String
    ToolDef_FreezePanes = TD("freeze_panes", _
        "Freeze or unfreeze rows/columns for scrolling. Freezes above and to the left of the specified cell.", _
        PS("cell", "Cell below and to the right of the freeze point e.g. A2 (freeze row 1), B3 (freeze row 1-2 and column A)") & "," & _
        PB("unfreeze", "Set true to remove all freeze panes") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        "")
End Function

Private Function ToolDef_RemoveDuplicates() As String
    ToolDef_RemoveDuplicates = TD("remove_duplicates", _
        "Remove duplicate rows from a range based on specified columns", _
        PS("range", "Range to deduplicate e.g. A1:D100") & "," & _
        PA("columns", "Column numbers (1-based within range) to check for duplicates e.g. [1,3]. Defaults to all columns.", "number") & "," & _
        PB("has_headers", "Whether first row contains headers (default: true)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range""")
End Function

Private Function ToolDef_SetValidation() As String
    ToolDef_SetValidation = TD("set_validation", _
        "Add data validation to cells (dropdown lists, number ranges, etc.)", _
        PS("range", "Range to validate e.g. C2:C100") & "," & _
        PS("type", "Validation type: list, whole_number, decimal, date, text_length") & "," & _
        PS("formula1", "For list: comma-separated values e.g. Yes,No,Maybe. For numbers: min value. For date: min date.") & "," & _
        PS("formula2", "For numbers/dates: max value (optional, for between operator)") & "," & _
        PS("operator", "Operator: between, not_between, equal, not_equal, greater, less, greater_equal, less_equal (default: between)") & "," & _
        PB("show_dropdown", "Show in-cell dropdown for list type (default: true)") & "," & _
        PS("error_message", "Custom error message (optional)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range"",""type"",""formula1""")
End Function

Private Function ToolDef_ListPivotTables() As String
    ToolDef_ListPivotTables = TD("list_pivot_tables", _
        "List all PivotTables in the workbook or a specific sheet with their names, source ranges and locations", _
        PS("sheet", "Sheet name (optional, lists all sheets if omitted)"), _
        "")
End Function

Private Function ToolDef_ClearRange() As String
    ToolDef_ClearRange = TD("clear_range", _
        "Clear contents, formatting, or everything from a range of cells", _
        PS("range", "Range to clear e.g. A1:D10") & "," & _
        PS("what", "What to clear: contents, formats, all (default: contents)") & "," & _
        PS("sheet", "Sheet name (optional)"), _
        """range""")
End Function
