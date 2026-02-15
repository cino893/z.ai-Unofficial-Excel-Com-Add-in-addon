Attribute VB_Name = "modJSON"
'==============================================================================
' modJSON - Lightweight JSON Parser/Builder for VBA
' Parses JSON into Dictionary/Collection, builds JSON from VBA objects
'==============================================================================
Option Explicit

Private p_pos As Long
Private p_json As String

' --- PUBLIC PARSE ---
Public Function JsonParse(ByVal jsonString As String) As Object
    p_json = jsonString
    p_pos = 1
    SkipWhitespace
    If Mid(p_json, p_pos, 1) = "{" Then
        Set JsonParse = ParseObject()
    ElseIf Mid(p_json, p_pos, 1) = "[" Then
        Set JsonParse = ParseArray()
    Else
        Set JsonParse = Nothing
    End If
End Function

' --- PUBLIC STRINGIFY ---
Public Function JsonStringify(ByVal obj As Variant, Optional ByVal indentLevel As Long = -1) As String
    If IsNull(obj) Or IsEmpty(obj) Then
        JsonStringify = "null"
    ElseIf IsObject(obj) Then
        If obj Is Nothing Then
            JsonStringify = "null"
        ElseIf TypeName(obj) = "Dictionary" Then
            JsonStringify = StringifyObject(obj, indentLevel)
        ElseIf TypeName(obj) = "Collection" Then
            JsonStringify = StringifyArray(obj, indentLevel)
        Else
            JsonStringify = "null"
        End If
    ElseIf VarType(obj) = vbString Then
        JsonStringify = """" & EscapeJsonString(CStr(obj)) & """"
    ElseIf VarType(obj) = vbBoolean Then
        JsonStringify = IIf(obj, "true", "false")
    ElseIf IsNumeric(obj) Then
        JsonStringify = CStr(obj)
        ' Ensure decimal point, not comma
        JsonStringify = Replace(JsonStringify, ",", ".")
    Else
        JsonStringify = """" & EscapeJsonString(CStr(obj)) & """"
    End If
End Function

' --- PARSE HELPERS ---
Private Function ParseObject() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    p_pos = p_pos + 1 ' skip {
    SkipWhitespace
    
    If Mid(p_json, p_pos, 1) = "}" Then
        p_pos = p_pos + 1
        Set ParseObject = dict
        Exit Function
    End If
    
    Do
        SkipWhitespace
        Dim key As String
        key = ParseString()
        SkipWhitespace
        
        If Mid(p_json, p_pos, 1) <> ":" Then
            Set ParseObject = dict
            Exit Function
        End If
        p_pos = p_pos + 1 ' skip :
        SkipWhitespace
        
        Dim value As Variant
        If IsNextValueObject() Then
            Set value = ParseValue()
            Set dict(key) = value
        Else
            value = ParseValue()
            If IsObject(value) Then
                Set dict(key) = value
            Else
                dict(key) = value
            End If
        End If
        
        SkipWhitespace
        If Mid(p_json, p_pos, 1) = "," Then
            p_pos = p_pos + 1
        ElseIf Mid(p_json, p_pos, 1) = "}" Then
            p_pos = p_pos + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop
    
    Set ParseObject = dict
End Function

Private Function ParseArray() As Object
    Dim coll As New Collection
    p_pos = p_pos + 1 ' skip [
    SkipWhitespace
    
    If Mid(p_json, p_pos, 1) = "]" Then
        p_pos = p_pos + 1
        Set ParseArray = coll
        Exit Function
    End If
    
    Do
        SkipWhitespace
        Dim value As Variant
        If IsNextValueObject() Then
            Set value = ParseValue()
            coll.Add value
        Else
            value = ParseValue()
            If IsObject(value) Then
                coll.Add value
            Else
                coll.Add value
            End If
        End If
        
        SkipWhitespace
        If Mid(p_json, p_pos, 1) = "," Then
            p_pos = p_pos + 1
        ElseIf Mid(p_json, p_pos, 1) = "]" Then
            p_pos = p_pos + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop
    
    Set ParseArray = coll
End Function

Private Function ParseValue() As Variant
    SkipWhitespace
    Dim ch As String
    ch = Mid(p_json, p_pos, 1)
    
    If ch = """" Then
        ParseValue = ParseString()
    ElseIf ch = "{" Then
        Set ParseValue = ParseObject()
    ElseIf ch = "[" Then
        Set ParseValue = ParseArray()
    ElseIf ch = "t" Then
        ' true
        p_pos = p_pos + 4
        ParseValue = True
    ElseIf ch = "f" Then
        ' false
        p_pos = p_pos + 5
        ParseValue = False
    ElseIf ch = "n" Then
        ' null
        p_pos = p_pos + 4
        ParseValue = Null
    Else
        ParseValue = ParseNumber()
    End If
End Function

Private Function ParseString() As String
    Dim result As String
    p_pos = p_pos + 1 ' skip opening "
    
    Do While p_pos <= Len(p_json)
        Dim ch As String
        ch = Mid(p_json, p_pos, 1)
        
        If ch = "\" Then
            p_pos = p_pos + 1
            Dim esc As String
            esc = Mid(p_json, p_pos, 1)
            Select Case esc
                Case """": result = result & """"
                Case "\": result = result & "\"
                Case "/": result = result & "/"
                Case "b": result = result & Chr(8)
                Case "f": result = result & Chr(12)
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "u"
                    Dim hex4 As String
                    hex4 = Mid(p_json, p_pos + 1, 4)
                    result = result & ChrW(CLng("&H" & hex4))
                    p_pos = p_pos + 4
            End Select
        ElseIf ch = """" Then
            p_pos = p_pos + 1
            Exit Do
        Else
            result = result & ch
        End If
        p_pos = p_pos + 1
    Loop
    
    ParseString = result
End Function

Private Function ParseNumber() As Variant
    Dim start As Long
    start = p_pos
    
    If Mid(p_json, p_pos, 1) = "-" Then p_pos = p_pos + 1
    
    Do While p_pos <= Len(p_json)
        Dim ch As String
        ch = Mid(p_json, p_pos, 1)
        If ch Like "[0-9]" Or ch = "." Or ch = "e" Or ch = "E" Or ch = "+" Or ch = "-" Then
            If start = p_pos And (ch = "+" Or ch = "-") Then Exit Do
            p_pos = p_pos + 1
        Else
            Exit Do
        End If
    Loop
    
    Dim numStr As String
    numStr = Mid(p_json, start, p_pos - start)
    
    If InStr(numStr, ".") > 0 Or InStr(numStr, "e") > 0 Or InStr(numStr, "E") > 0 Then
        ParseNumber = CDbl(Replace(numStr, ".", Application.International(xlDecimalSeparator)))
    Else
        If CDbl(numStr) > 2147483647# Or CDbl(numStr) < -2147483648# Then
            ParseNumber = CDbl(numStr)
        Else
            ParseNumber = CLng(numStr)
        End If
    End If
End Function

Private Function IsNextValueObject() As Boolean
    SkipWhitespace
    Dim ch As String
    ch = Mid(p_json, p_pos, 1)
    IsNextValueObject = (ch = "{" Or ch = "[")
End Function

Private Sub SkipWhitespace()
    Do While p_pos <= Len(p_json)
        Select Case Mid(p_json, p_pos, 1)
            Case " ", vbTab, vbCr, vbLf
                p_pos = p_pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

' --- STRINGIFY HELPERS ---
Private Function StringifyObject(ByVal dict As Object, ByVal indentLevel As Long) As String
    Dim result As String
    Dim keys() As Variant
    Dim i As Long
    
    result = "{"
    keys = dict.keys
    
    For i = 0 To UBound(keys)
        If i > 0 Then result = result & ","
        result = result & """" & EscapeJsonString(CStr(keys(i))) & """:"
        result = result & JsonStringify(dict(keys(i)), indentLevel)
    Next i
    
    result = result & "}"
    StringifyObject = result
End Function

Private Function StringifyArray(ByVal coll As Collection, ByVal indentLevel As Long) As String
    Dim result As String
    Dim i As Long
    
    result = "["
    For i = 1 To coll.Count
        If i > 1 Then result = result & ","
        result = result & JsonStringify(coll(i), indentLevel)
    Next i
    
    result = result & "]"
    StringifyArray = result
End Function

Public Function EscapeJsonString(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid(s, i, 1)
        Select Case ch
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case vbLf
                result = result & "\n"
            Case vbCr
                result = result & "\r"
            Case vbTab
                result = result & "\t"
            Case Else
                If AscW(ch) < 32 Then
                    result = result & "\u" & Right("0000" & Hex(AscW(ch)), 4)
                Else
                    result = result & ch
                End If
        End Select
    Next i
    
    EscapeJsonString = result
End Function

' --- UTILITY: Create Dictionary ---
Public Function JsonDict(ParamArray keyValuePairs() As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        If i + 1 <= UBound(keyValuePairs) Then
            If IsObject(keyValuePairs(i + 1)) Then
                Set dict(CStr(keyValuePairs(i))) = keyValuePairs(i + 1)
            Else
                dict(CStr(keyValuePairs(i))) = keyValuePairs(i + 1)
            End If
        End If
    Next i
    Set JsonDict = dict
End Function

' --- UTILITY: Create Collection (Array) ---
Public Function JsonArr(ParamArray items() As Variant) As Collection
    Dim coll As New Collection
    Dim i As Long
    For i = LBound(items) To UBound(items)
        coll.Add items(i)
    Next i
    Set JsonArr = coll
End Function

' --- UTILITY: Safe dictionary access ---
Public Function DictGet(ByVal dict As Object, ByVal key As String, Optional ByVal defaultVal As Variant = "") As Variant
    If dict Is Nothing Then
        DictGet = defaultVal
        Exit Function
    End If
    If dict.Exists(key) Then
        If IsObject(dict(key)) Then
            Set DictGet = dict(key)
        Else
            DictGet = dict(key)
        End If
    Else
        DictGet = defaultVal
    End If
End Function
