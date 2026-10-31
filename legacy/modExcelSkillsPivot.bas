Attribute VB_Name = "modExcelSkillsPivot"
'==============================================================================
' modExcelSkillsPivot - PivotTable skills matching COM ExcelSkillService.Charts
'==============================================================================
Option Explicit

Private Const XL_DATABASE As Long = 1
Private Const XL_ROW_FIELD As Long = 1
Private Const XL_COLUMN_FIELD As Long = 2
Private Const XL_PAGE_FIELD As Long = 3
Private Const XL_DATA_FIELD As Long = 4
Private Const XL_HIDDEN As Long = 0
Private Const TYPE_E_INVDATAREAD As Long = -2147319784

' --- create_pivot_table ---
Public Function SkillCreatePivotTable(ByVal args As Object) As String
    Dim app As Application
    Set app = Application
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    Dim wb As Workbook
    Set wb = app.ActiveWorkbook
    
    Dim sourceRange As String, destCell As String, ptName As String
    sourceRange = ArgStr(args, "source_range")
    destCell = ArgStr(args, "dest_cell")
    ptName = ArgStr(args, "name")
    If ptName = "" Then ptName = "PivotTable" & (wb.PivotCaches.Count + 1)
    
    Dim srcRange As Range
    Set srcRange = ws.Range(sourceRange)
    Dim qualifiedSource As String
    qualifiedSource = "'" & ws.Name & "'!" & srcRange.Address
    
    Dim cache As PivotCache
    Set cache = wb.PivotCaches.Create(SourceType:=XL_DATABASE, SourceData:=qualifiedSource)
    
    Dim destSheet As Worksheet
    Dim destRange As Range
    Dim destSheetName As String
    destSheetName = ArgStr(args, "dest_sheet")
    
    If destSheetName <> "" Or destCell = "" Then
        Set destSheet = ResolveDestSheet(wb, destSheetName)
        If destCell = "" Then
            Set destRange = destSheet.Range("A3")
        Else
            Set destRange = destSheet.Range(destCell)
        End If
    Else
        Set destSheet = ws
        Set destRange = destSheet.Range(destCell)
    End If
    
    Dim pt As PivotTable
    Set pt = cache.CreatePivotTable(TableDestination:=destRange, TableName:=ptName)
    
    Dim rowFields As Collection, colFields As Collection, valFields As Collection
    Set rowFields = ArgColl(args, "row_fields")
    Set colFields = ArgColl(args, "column_fields")
    Set valFields = ArgColl(args, "value_fields")
    
    Dim i As Long
    If Not rowFields Is Nothing Then
        For i = 1 To rowFields.Count
            pt.PivotFields(CStr(rowFields(i))).Orientation = XL_ROW_FIELD
        Next i
    End If
    If Not colFields Is Nothing Then
        For i = 1 To colFields.Count
            pt.PivotFields(CStr(colFields(i))).Orientation = XL_COLUMN_FIELD
        Next i
    End If
    
    Dim xlFunc As Long
    xlFunc = XlValueFunction(ArgStr(args, "value_function", "sum"))
    If Not valFields Is Nothing Then
        For i = 1 To valFields.Count
            Dim pf As PivotField
            Set pf = pt.PivotFields(CStr(valFields(i)))
            pf.Orientation = XL_DATA_FIELD
            pf.Function = xlFunc
        Next i
    End If
    
    Dim rowCount As Long, valCount As Long
    If Not rowFields Is Nothing Then rowCount = rowFields.Count
    If Not valFields Is Nothing Then valCount = valFields.Count
    
    SkillCreatePivotTable = "{""success"":true," & JsonStr("name", ptName) & "," & _
        JsonStr("dest_sheet", destSheet.Name) & "," & _
        JsonNum("row_fields", rowCount) & "," & JsonNum("value_fields", valCount) & "}"
End Function

' --- list_pivot_tables ---
Public Function SkillListPivotTables(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim sheetFilter As String
    sheetFilter = ArgStr(args, "sheet")
    
    Dim pivotsJson As String
    pivotsJson = "["
    Dim pivotCount As Long
    pivotCount = 0
    Dim errorsJson As String
    errorsJson = "["
    Dim errCount As Long
    errCount = 0
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If sheetFilter = "" Or LCase$(ws.Name) = LCase$(sheetFilter) Then
            On Error Resume Next
            Dim pt As PivotTable
            Err.Clear
            For Each pt In ws.PivotTables
                If Err.Number <> 0 Then
                    errCount = errCount + 1
                    If errCount > 1 Then errorsJson = errorsJson & ","
                    errorsJson = errorsJson & """" & EscapeJsonString(ws.Name) & """"
                    Err.Clear
                    Exit For
                End If
                Dim location As String, source As String
                location = ""
                source = ""
                On Error Resume Next
                location = pt.TableRange2.Address
                source = CStr(pt.SourceData)
                On Error GoTo 0
                If pivotCount > 0 Then pivotsJson = pivotsJson & ","
                pivotsJson = pivotsJson & "{" & JsonStr("name", CStr(pt.Name)) & "," & _
                    JsonStr("sheet", ws.Name) & "," & JsonStr("location", location) & "," & _
                    JsonStr("source", source) & "}"
                pivotCount = pivotCount + 1
                On Error Resume Next
            Next pt
            If Err.Number <> 0 Then
                Dim already As Boolean
                already = False
                If InStr(errorsJson, """" & EscapeJsonString(ws.Name) & """") > 0 Then already = True
                If Not already Then
                    errCount = errCount + 1
                    If errCount > 1 Then errorsJson = errorsJson & ","
                    errorsJson = errorsJson & """" & EscapeJsonString(ws.Name) & """"
                End If
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next ws
    pivotsJson = pivotsJson & "]"
    errorsJson = errorsJson & "]"
    
    Dim result As String
    result = "{""pivot_tables"":" & pivotsJson & "," & JsonNum("count", pivotCount)
    If errCount > 0 Then
        result = result & ",""errors"":" & errorsJson
        Dim cacheJson As String
        cacheJson = GetPivotCacheSourcesJson(wb)
        If cacheJson <> "[]" Then
            result = result & ",""pivot_cache_sources"":" & cacheJson
        End If
        Dim srcInfo As String
        srcInfo = ""
        If cacheJson <> "[]" Then
            Dim firstSrc As String
            firstSrc = FirstJsonArrayItem(cacheJson)
            If firstSrc <> "" Then srcInfo = " The pivot data source is: " & firstSrc & "."
        End If
        result = result & "," & JsonStr("hint", _
            "PivotTable enumeration failed (COM type library issue, common with .xlsb files). " & _
            "Do NOT retry list_pivot_tables or move_pivot_table." & srcInfo & _
            " To recreate a pivot on another sheet: call create_pivot_table with the source data range, then modify_pivot_table to configure fields.")
    End If
    result = result & "}"
    SkillListPivotTables = result
End Function

' --- move_pivot_table ---
Public Function SkillMovePivotTable(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim sourceName As String, sourceRange As String, destSheetName As String, destCell As String
    sourceName = ArgStr(args, "name")
    sourceRange = ArgStr(args, "source_range")
    destSheetName = ArgStr(args, "dest_sheet")
    destCell = ArgStr(args, "dest_cell", "A1")
    Dim sourceSheet As Worksheet
    Set sourceSheet = GetTargetSheet(args)
    
    Dim foundPivot As PivotTable
    Dim pivotSheet As Worksheet
    Dim comTypeLibError As Boolean
    comTypeLibError = False
    
    Dim sheet As Worksheet
    Dim pt As PivotTable
    
    ' TIER 1: PivotTables enumeration by name
    If sourceName <> "" Then
        For Each sheet In wb.Worksheets
            On Error Resume Next
            Err.Clear
            For Each pt In sheet.PivotTables
                If Err.Number = TYPE_E_INVDATAREAD Then
                    comTypeLibError = True
                    Err.Clear
                    Exit For
                ElseIf Err.Number <> 0 Then
                    Err.Clear
                    Exit For
                End If
                If CStr(pt.Name) = sourceName Then
                    Set foundPivot = pt
                    Set pivotSheet = sheet
                    Exit For
                End If
            Next pt
            On Error GoTo 0
            If Not foundPivot Is Nothing Then Exit For
        Next sheet
    End If
    
    ' TIER 2: Direct name lookup
    If foundPivot Is Nothing And sourceName <> "" Then
        For Each sheet In wb.Worksheets
            On Error Resume Next
            Err.Clear
            Set pt = sheet.PivotTables(sourceName)
            If Err.Number = TYPE_E_INVDATAREAD Then
                comTypeLibError = True
                Err.Clear
            ElseIf Err.Number = 0 Then
                If Not pt Is Nothing Then
                    Set foundPivot = pt
                    Set pivotSheet = sheet
                    On Error GoTo 0
                    Exit For
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0
        Next sheet
    End If
    
    ' TIER 3: Range.PivotTable
    If foundPivot Is Nothing And sourceRange <> "" Then
        On Error Resume Next
        Err.Clear
        Set foundPivot = sourceSheet.Range(sourceRange).Cells(1, 1).PivotTable
        If Err.Number = 0 Then
            Set pivotSheet = sourceSheet
        Else
            If Err.Number = TYPE_E_INVDATAREAD Then comTypeLibError = True
            Set foundPivot = Nothing
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' TIER 3b: Intersection check
    If foundPivot Is Nothing And sourceRange <> "" And Not comTypeLibError Then
        On Error Resume Next
        Dim srcRange As Range
        Set srcRange = sourceSheet.Range(sourceRange)
        For Each pt In sourceSheet.PivotTables
            Dim ptRange As Range
            Set ptRange = pt.TableRange2
            Dim overlap As Range
            Set overlap = Application.Intersect(srcRange, ptRange)
            If Not overlap Is Nothing Then
                Set foundPivot = pt
                Set pivotSheet = sourceSheet
                Exit For
            End If
        Next pt
        On Error GoTo 0
    End If
    
    ' Safety net: PivotCaches exist
    If foundPivot Is Nothing And Not comTypeLibError Then
        On Error Resume Next
        If wb.PivotCaches.Count > 0 Then comTypeLibError = True
        On Error GoTo 0
    End If
    
    LogDebug "move_pivot_table: name=" & sourceName & ", range=" & sourceRange & _
        ", found=" & (Not foundPivot Is Nothing) & ", comErr=" & comTypeLibError
    
    If foundPivot Is Nothing And Not comTypeLibError Then
        SkillMovePivotTable = "{""error"":""No pivot table found. Check the name or source_range.""," & _
            """hint"":""Use list_pivot_tables to find pivot tables, or use copy_range to move plain data.""}"
        Exit Function
    End If
    
    Dim destSheet As Worksheet
    Set destSheet = ResolveDestSheet(wb, destSheetName)
    Dim destRng As Range
    Set destRng = destSheet.Range(destCell)
    
    ' A. Location move
    If Not foundPivot Is Nothing Then
        Dim fromSheet As String, ptName As String, destSheetActual As String
        fromSheet = pivotSheet.Name
        ptName = CStr(foundPivot.Name)
        destSheetActual = destSheet.Name
        On Error Resume Next
        Err.Clear
        foundPivot.Location = "'" & destSheetActual & "'!" & destCell
        If Err.Number = 0 Then
            On Error GoTo 0
            SkillMovePivotTable = "{""success"":true,""method"":""location_move""," & _
                JsonStr("name", ptName) & "," & JsonStr("from_sheet", fromSheet) & "," & _
                JsonStr("to_sheet", destSheetActual) & "," & JsonStr("dest_cell", destCell) & "}"
            Exit Function
        Else
            LogDebug "move_pivot_table Location failed: " & Err.Number & " " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' B. Recreate from PivotCaches
    On Error Resume Next
    Dim cache As PivotCache
    For Each cache In wb.PivotCaches
        Err.Clear
        Dim sd As String
        sd = ""
        sd = CStr(cache.SourceData)
        If sd <> "" Then
            Dim newName As String
            If sourceName <> "" Then newName = sourceName & "_moved" Else newName = "PivotTable_moved"
            cache.CreatePivotTable TableDestination:=destRng, TableName:=newName
            If Err.Number = 0 Then
                Dim clearNote As String
                clearNote = ""
                If sourceRange <> "" Then
                    Err.Clear
                    sourceSheet.Range(sourceRange).Clear
                    If Err.Number = 0 Then
                        clearNote = " Old range cleared."
                    Else
                        clearNote = " Clear old range manually with clear_range."
                        Err.Clear
                    End If
                End If
                On Error GoTo 0
                SkillMovePivotTable = "{""success"":true,""method"":""pivot_recreated""," & _
                    JsonStr("new_name", newName) & "," & JsonStr("source_data", ConvertR1C1toA1(sd)) & "," & _
                    JsonStr("to_sheet", destSheet.Name) & "," & JsonStr("dest_cell", destCell) & "," & _
                    JsonStr("next_step", "Pivot recreated empty — use modify_pivot_table to configure row_fields, value_fields, column_fields." & clearNote) & "}"
                Exit Function
            End If
            Err.Clear
        End If
    Next cache
    On Error GoTo 0
    
    ' C. All methods failed
    Dim cacheSources As String
    cacheSources = GetPivotCacheSourcesJson(wb)
    Dim srcHint As String
    srcHint = ""
    If cacheSources <> "[]" Then srcHint = " (source_range: " & FirstJsonArrayItem(cacheSources) & ")"
    Dim destHint As String
    If destSheetName = "" Then destHint = "new sheet" Else destHint = destSheetName
    SkillMovePivotTable = "{""error"":""Cannot move pivot table — COM access denied on this workbook.""," & _
        """alternative"":""Use create_pivot_table with source data" & EscapeJsonString(srcHint) & _
        " on dest_sheet='" & EscapeJsonString(destHint) & "', then clear_range on old pivot area.""}"
End Function

' --- modify_pivot_table ---
Public Function SkillModifyPivotTable(ByVal args As Object) As String
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim ptName As String
    ptName = ArgStr(args, "name")
    Dim ws As Worksheet
    Set ws = GetTargetSheet(args)
    
    Dim pt As PivotTable
    If ptName <> "" Then
        On Error Resume Next
        Set pt = ws.PivotTables(ptName)
        If pt Is Nothing Then
            Dim sheet As Worksheet
            For Each sheet In wb.Worksheets
                Err.Clear
                Set pt = sheet.PivotTables(ptName)
                If Not pt Is Nothing Then Exit For
            Next sheet
        End If
        On Error GoTo 0
    Else
        On Error Resume Next
        If ws.PivotTables.Count > 0 Then Set pt = ws.PivotTables(1)
        On Error GoTo 0
    End If
    
    If pt Is Nothing Then
        SkillModifyPivotTable = "{""error"":""Pivot table '" & EscapeJsonString(ptName) & _
            "' not found. Use list_pivot_tables to find available pivots.""}"
        Exit Function
    End If
    
    Dim changes As Long
    changes = 0
    Dim hasNewFields As Boolean
    hasNewFields = ArgHas(args, "row_fields") Or ArgHas(args, "column_fields") Or _
                   ArgHas(args, "value_fields") Or ArgHas(args, "page_fields")
    
    If hasNewFields And ArgBool(args, "clear_fields", True) Then
        On Error Resume Next
        Dim pf As PivotField
        For Each pf In pt.PivotFields
            If pf.Orientation <> XL_HIDDEN Then pf.Orientation = XL_HIDDEN
        Next pf
        On Error GoTo 0
    End If
    
    Dim i As Long
    Dim fields As Collection
    Set fields = ArgColl(args, "row_fields")
    If Not fields Is Nothing Then
        For i = 1 To fields.Count
            On Error Resume Next
            Set pf = pt.PivotFields(CStr(fields(i)))
            If Err.Number = 0 Then
                pf.Orientation = XL_ROW_FIELD
                changes = changes + 1
            Else
                LogDebug "modify_pivot: row field '" & CStr(fields(i)) & "' failed: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Next i
    End If
    
    Set fields = ArgColl(args, "column_fields")
    If Not fields Is Nothing Then
        For i = 1 To fields.Count
            On Error Resume Next
            Set pf = pt.PivotFields(CStr(fields(i)))
            If Err.Number = 0 Then
                pf.Orientation = XL_COLUMN_FIELD
                changes = changes + 1
            Else
                LogDebug "modify_pivot: col field '" & CStr(fields(i)) & "' failed: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Next i
    End If
    
    Set fields = ArgColl(args, "value_fields")
    If Not fields Is Nothing Then
        Dim xlFunc As Long
        xlFunc = XlValueFunction(ArgStr(args, "value_function", "sum"))
        For i = 1 To fields.Count
            On Error Resume Next
            Set pf = pt.PivotFields(CStr(fields(i)))
            If Err.Number = 0 Then
                pf.Orientation = XL_DATA_FIELD
                pf.Function = xlFunc
                changes = changes + 1
            Else
                LogDebug "modify_pivot: value field '" & CStr(fields(i)) & "' failed: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Next i
    End If
    
    Set fields = ArgColl(args, "page_fields")
    If Not fields Is Nothing Then
        For i = 1 To fields.Count
            On Error Resume Next
            Set pf = pt.PivotFields(CStr(fields(i)))
            If Err.Number = 0 Then
                pf.Orientation = XL_PAGE_FIELD
                changes = changes + 1
            Else
                LogDebug "modify_pivot: page field '" & CStr(fields(i)) & "' failed: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Next i
    End If
    
    If ArgBool(args, "refresh") Then
        On Error Resume Next
        pt.RefreshTable
        On Error GoTo 0
    End If
    
    Dim rowsJ As String, colsJ As String, valsJ As String, pagesJ As String
    rowsJ = "[": colsJ = "[": valsJ = "[": pagesJ = "["
    Dim rc As Long, cc As Long, vc As Long, pc As Long
    On Error Resume Next
    For Each pf In pt.PivotFields
        Dim fn As String, o As Long
        fn = CStr(pf.Name)
        o = CLng(pf.Orientation)
        If o = XL_ROW_FIELD Then
            If rc > 0 Then rowsJ = rowsJ & ","
            rowsJ = rowsJ & """" & EscapeJsonString(fn) & """"
            rc = rc + 1
        ElseIf o = XL_COLUMN_FIELD Then
            If cc > 0 Then colsJ = colsJ & ","
            colsJ = colsJ & """" & EscapeJsonString(fn) & """"
            cc = cc + 1
        ElseIf o = XL_DATA_FIELD Then
            If vc > 0 Then valsJ = valsJ & ","
            valsJ = valsJ & """" & EscapeJsonString(fn) & """"
            vc = vc + 1
        ElseIf o = XL_PAGE_FIELD Then
            If pc > 0 Then pagesJ = pagesJ & ","
            pagesJ = pagesJ & """" & EscapeJsonString(fn) & """"
            pc = pc + 1
        End If
    Next pf
    On Error GoTo 0
    rowsJ = rowsJ & "]": colsJ = colsJ & "]": valsJ = valsJ & "]": pagesJ = pagesJ & "]"
    
    Dim layout As String
    layout = "{""row_fields"":" & rowsJ & ",""column_fields"":" & colsJ & ",""value_fields"":" & valsJ
    If pc > 0 Then layout = layout & ",""page_fields"":" & pagesJ
    layout = layout & "}"
    
    SkillModifyPivotTable = "{""success"":true," & JsonStr("name", CStr(pt.Name)) & "," & _
        JsonNum("changes", changes) & ",""current_layout"":" & layout & "}"
End Function

' --- helpers ---
Public Function ResolveDestSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If sheetName = "" Then
        Set ResolveDestSheet = wb.Worksheets.Add
        Exit Function
    End If
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        On Error Resume Next
        ws.Name = sheetName
        On Error GoTo 0
    End If
    Set ResolveDestSheet = ws
End Function

Public Function ConvertR1C1toA1(ByVal r1c1 As String) As String
    Dim bangIdx As Long
    bangIdx = InStrRev(r1c1, "!")
    Dim sheet As String, refPart As String
    If bangIdx > 0 Then
        sheet = Left$(r1c1, bangIdx)
        refPart = Mid$(r1c1, bangIdx + 1)
    Else
        sheet = ""
        refPart = r1c1
    End If
    
    On Error Resume Next
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[WRwr](\d+)[KCkc](\d+)(?::?[WRwr](\d+)[KCkc](\d+))?$"
    re.IgnoreCase = True
    Dim matches As Object
    Set matches = re.Execute(refPart)
    If matches Is Nothing Or matches.Count = 0 Then
        ConvertR1C1toA1 = r1c1
        Exit Function
    End If
    Dim g As Object
    Set g = matches(0).SubMatches
    Dim r1 As Long, c1 As Long
    r1 = CLng(g(0))
    c1 = CLng(g(1))
    Dim result As String
    result = sheet & ColNumToLetter(c1) & CStr(r1)
    If g.Count >= 4 Then
        If CStr(g(2)) <> "" Then
            Dim r2 As Long, c2 As Long
            r2 = CLng(g(2))
            c2 = CLng(g(3))
            result = result & ":" & ColNumToLetter(c2) & CStr(r2)
        End If
    End If
    ConvertR1C1toA1 = result
    On Error GoTo 0
End Function

Public Function GetPivotCacheSourcesJson(ByVal wb As Workbook) As String
    Dim json As String
    json = "["
    Dim n As Long
    n = 0
    On Error Resume Next
    Dim cache As PivotCache
    For Each cache In wb.PivotCaches
        Dim sd As String
        sd = ""
        sd = CStr(cache.SourceData)
        If sd <> "" Then
            If n > 0 Then json = json & ","
            json = json & """" & EscapeJsonString(ConvertR1C1toA1(sd)) & """"
            n = n + 1
        End If
    Next cache
    On Error GoTo 0
    json = json & "]"
    GetPivotCacheSourcesJson = json
End Function

Private Function FirstJsonArrayItem(ByVal arrJson As String) As String
    ' arrJson like ["Foo"] or ["Foo","Bar"] — return first unquoted-ish value
    Dim s As String
    s = Trim$(arrJson)
    If Left$(s, 1) = "[" Then s = Mid$(s, 2)
    If Right$(s, 1) = "]" Then s = Left$(s, Len(s) - 1)
    s = Trim$(s)
    If Left$(s, 1) = """" Then
        Dim endQ As Long
        endQ = InStr(2, s, """")
        If endQ > 0 Then
            FirstJsonArrayItem = Mid$(s, 2, endQ - 2)
            Exit Function
        End If
    End If
    FirstJsonArrayItem = s
End Function
