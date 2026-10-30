Attribute VB_Name = "modZaiAPI"
'==============================================================================
' modZaiAPI - z.ai API Communication Module
' Mirrors src/Services/ZaiApiService.cs
'==============================================================================
Option Explicit

Public Const ZAI_API_BASE As String = "https://api.z.ai/api/paas/v4"
Public Const ZAI_DEFAULT_MODEL As String = "glm-4.7-flash"
Public Const ZAI_MAX_TOKENS As Long = 4096
Public Const ZAI_TEMPERATURE As Double = 0.7
Private Const HTTP_TIMEOUT_SEC As Long = 120

' --- Send Chat Completion Request ---
Public Function SendChatCompletion( _
    ByVal messagesJson As String, _
    Optional ByVal toolsJson As String = "", _
    Optional ByVal model As String = "" _
) As Object
    
    If model = "" Then model = LoadModel()
    
    Dim apiKey As String
    apiKey = SanitizeApiKey(LoadApiKey())
    If apiKey = "" Then
        LogError "SendChatCompletion: No API key"
        Dim noKey As Object
        Set noKey = CreateObject("Scripting.Dictionary")
        noKey("success") = False
        noKey("error") = "No API key configured"
        Set SendChatCompletion = noKey
        Exit Function
    End If
    
    Dim body As String
    body = "{" & _
           """model"":""" & EscapeJsonString(model) & """," & _
           """messages"":" & messagesJson & "," & _
           """max_tokens"":" & ZAI_MAX_TOKENS & "," & _
           """temperature"":" & Replace(CStr(ZAI_TEMPERATURE), ",", ".")
    
    If toolsJson <> "" Then
        body = body & ",""tools"":" & toolsJson & ",""tool_choice"":""auto"""
    End If
    body = body & "}"
    
    Dim url As String
    url = ZAI_API_BASE & "/chat/completions"
    
    Set SendChatCompletion = HttpPost(url, body, apiKey)
End Function

' --- HTTP POST (async with cancel, sync fallback) ---
Public Function HttpPost(ByVal url As String, ByVal body As String, ByVal apiKey As String) As Object
    On Error GoTo ErrHandler
    
    apiKey = SanitizeApiKey(apiKey)
    LogApiRequest "POST", url, body
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim usedAsync As Boolean
    usedAsync = False
    On Error Resume Next
    Err.Clear
    http.Open "POST", url, True
    If Err.Number = 0 Then
        usedAsync = True
    Else
        Err.Clear
        http.Open "POST", url, False
    End If
    On Error GoTo ErrHandler
    
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.setRequestHeader "Accept", "application/json"
    http.send body
    
    If usedAsync Then
        Dim start As Single
        start = Timer
        Do While http.readyState <> 4
            If IsCancelRequested() Then
                On Error Resume Next
                http.abort
                On Error GoTo 0
                Dim cancelled As Object
                Set cancelled = CreateObject("Scripting.Dictionary")
                cancelled("success") = False
                cancelled("cancelled") = True
                cancelled("status") = 0
                cancelled("error") = ""
                LogInfo "HTTP request cancelled"
                Set HttpPost = cancelled
                Exit Function
            End If
            Dim elapsed As Single
            elapsed = Timer - start
            If elapsed < 0 Then elapsed = elapsed + 86400
            If elapsed > HTTP_TIMEOUT_SEC Then
                On Error Resume Next
                http.abort
                On Error GoTo 0
                Dim toRes As Object
                Set toRes = CreateObject("Scripting.Dictionary")
                toRes("success") = False
                toRes("status") = 0
                toRes("error") = "Network error: timeout"
                Set HttpPost = toRes
                Exit Function
            End If
            DoEvents
        Loop
    End If
    
    LogApiResponse http.Status, http.responseText
    
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("status") = http.Status
    result("raw") = http.responseText
    
    If http.Status = 200 Then
        Dim parsed As Object
        Set parsed = JsonParse(http.responseText)
        If Not parsed Is Nothing Then Set result("data") = parsed
        result("success") = True
    Else
        result("success") = False
        result("error") = TranslateApiError(CLng(http.Status), http.responseText)
        LogError "API error: HTTP " & http.Status
    End If
    
    Set HttpPost = result
    Exit Function
    
ErrHandler:
    LogErrorDetails "HttpPost", Err.Number, Err.Description
    Dim errResult As Object
    Set errResult = CreateObject("Scripting.Dictionary")
    errResult("success") = False
    errResult("status") = 0
    errResult("error") = "Network error: " & Err.Description
    Set HttpPost = errResult
End Function

Public Function TranslateApiError(ByVal httpCode As Long, ByVal body As String) As String
    On Error GoTo Fallback
    Dim json As Object
    Set json = JsonParse(body)
    Dim code As String, message As String
    code = ""
    message = ""
    If Not json Is Nothing Then
        If json.Exists("error") Then
            If IsObject(json("error")) Then
                Dim errObj As Object
                Set errObj = json("error")
                code = CStr(DictGet(errObj, "code", ""))
                message = CStr(DictGet(errObj, "message", ""))
            End If
        End If
    End If
    
    Select Case code
        Case "1113", "1261"
            TranslateApiError = T("error.balance_empty")
            Exit Function
        Case "1301"
            TranslateApiError = T("error.content_filter")
            Exit Function
        Case "1302"
            TranslateApiError = T("error.rate_limit")
            Exit Function
    End Select
    
    Dim msgL As String
    msgL = LCase$(message)
    If InStr(msgL, "balance") > 0 Or InStr(msgL, "recharge") > 0 Or _
       InStr(message, ChrW(&H4F59) & ChrW(&H989D)) > 0 Or _
       InStr(message, ChrW(&H5145) & ChrW(&H503C)) > 0 Then
        TranslateApiError = T("error.balance_empty")
        Exit Function
    End If
    
Fallback:
    Select Case httpCode
        Case 401: TranslateApiError = T("error.invalid_key")
        Case 429: TranslateApiError = T("error.rate_limit")
        Case Else: TranslateApiError = T("error.api_generic") & " (HTTP " & httpCode & ")"
    End Select
End Function

' --- Model catalog (id | input | output | sort) ---
Public Function GetModelCatalogCount() As Long
    GetModelCatalogCount = 12
End Function

Public Function GetModelCatalogId(ByVal index As Long) As String
    Select Case index
        Case 1: GetModelCatalogId = "glm-4.7-flash"
        Case 2: GetModelCatalogId = "glm-4.5-flash"
        Case 3: GetModelCatalogId = "glm-4.7-flashx"
        Case 4: GetModelCatalogId = "glm-4-32b-0414-128k"
        Case 5: GetModelCatalogId = "glm-4.5-air"
        Case 6: GetModelCatalogId = "glm-4.5"
        Case 7: GetModelCatalogId = "glm-4.6"
        Case 8: GetModelCatalogId = "glm-4.7"
        Case 9: GetModelCatalogId = "glm-4.5-airx"
        Case 10: GetModelCatalogId = "glm-5"
        Case 11: GetModelCatalogId = "glm-5-code"
        Case 12: GetModelCatalogId = "glm-4.5-x"
        Case Else: GetModelCatalogId = ZAI_DEFAULT_MODEL
    End Select
End Function

Public Function GetModelCatalogDisplay(ByVal index As Long) As String
    Dim id As String
    id = GetModelCatalogId(index)
    If index <= 2 Then
        GetModelCatalogDisplay = "*  " & id & "   —   FREE"
    ElseIf index < 6 Then
        GetModelCatalogDisplay = "+  " & id & "   —   " & PriceTag(index)
    ElseIf index < 10 Then
        GetModelCatalogDisplay = "#  " & id & "   —   " & PriceTag(index)
    Else
        GetModelCatalogDisplay = "$  " & id & "   —   " & PriceTag(index)
    End If
End Function

Private Function PriceTag(ByVal index As Long) As String
    Select Case index
        Case 1, 2: PriceTag = "FREE"
        Case 3: PriceTag = "$0.07→$0.4/MTok"
        Case 4: PriceTag = "$0.1→$0.1/MTok"
        Case 5: PriceTag = "$0.2→$1.1/MTok"
        Case 6: PriceTag = "$0.6→$2.2/MTok"
        Case 7: PriceTag = "$0.6→$2.2/MTok"
        Case 8: PriceTag = "$0.6→$2.2/MTok"
        Case 9: PriceTag = "$1.1→$4.5/MTok"
        Case 10: PriceTag = "$1→$3.2/MTok"
        Case 11: PriceTag = "$1.2→$5/MTok"
        Case 12: PriceTag = "$2.2→$8.9/MTok"
        Case Else: PriceTag = "unknown"
    End Select
End Function

Public Function GetModelsPromptList() As String
    Dim i As Long, result As String
    result = ""
    For i = 1 To GetModelCatalogCount()
        If result <> "" Then result = result & vbCrLf
        result = result & GetModelCatalogDisplay(i)
    Next i
    GetModelsPromptList = result
End Function

Public Function GetResponseContent(ByVal response As Object) As String
    On Error GoTo ErrHandler
    If response Is Nothing Then
        GetResponseContent = ""
        Exit Function
    End If
    If Not response("success") Then
        GetResponseContent = DictGet(response, "error", T("conv.no_response"))
        Exit Function
    End If
    Dim data As Object
    Set data = response("data")
    Dim choices As Object
    Set choices = data("choices")
    If choices.Count = 0 Then
        GetResponseContent = ""
        Exit Function
    End If
    Dim firstChoice As Object
    Set firstChoice = choices(1)
    Dim msg As Object
    Set msg = firstChoice("message")
    If msg.Exists("content") Then
        If Not IsNull(msg("content")) Then GetResponseContent = CStr(msg("content"))
    End If
    Exit Function
ErrHandler:
    LogErrorDetails "GetResponseContent", Err.Number, Err.Description
    GetResponseContent = ""
End Function

Public Function GetToolCalls(ByVal response As Object) As Collection
    Dim result As New Collection
    On Error GoTo ErrHandler
    If response Is Nothing Then
        Set GetToolCalls = result
        Exit Function
    End If
    If Not response("success") Then
        Set GetToolCalls = result
        Exit Function
    End If
    Dim data As Object
    Set data = response("data")
    Dim choices As Object
    Set choices = data("choices")
    If choices.Count = 0 Then
        Set GetToolCalls = result
        Exit Function
    End If
    Dim firstChoice As Object
    Set firstChoice = choices(1)
    Dim msg As Object
    Set msg = firstChoice("message")
    If msg.Exists("tool_calls") Then
        If IsObject(msg("tool_calls")) Then
            Set result = msg("tool_calls")
            LogInfo "Found " & result.Count & " tool call(s) in response"
        End If
    End If
    Set GetToolCalls = result
    Exit Function
ErrHandler:
    LogErrorDetails "GetToolCalls", Err.Number, Err.Description
    Set GetToolCalls = New Collection
End Function

Public Function HasToolCalls(ByVal response As Object) As Boolean
    On Error GoTo ErrHandler
    Dim calls As Collection
    Set calls = GetToolCalls(response)
    HasToolCalls = (calls.Count > 0)
    Exit Function
ErrHandler:
    HasToolCalls = False
End Function

Public Function GetFinishReason(ByVal response As Object) As String
    On Error GoTo ErrHandler
    If response Is Nothing Or Not response("success") Then
        GetFinishReason = "error"
        Exit Function
    End If
    Dim data As Object
    Set data = response("data")
    Dim choices As Object
    Set choices = data("choices")
    If choices.Count > 0 Then
        Dim firstChoice As Object
        Set firstChoice = choices(1)
        GetFinishReason = DictGet(firstChoice, "finish_reason", "unknown")
    Else
        GetFinishReason = "no_choices"
    End If
    Exit Function
ErrHandler:
    GetFinishReason = "error"
End Function

Public Function GetAssistantMessage(ByVal response As Object) As Object
    On Error GoTo ErrHandler
    Dim data As Object
    Set data = response("data")
    Dim choices As Object
    Set choices = data("choices")
    If choices.Count > 0 Then
        Dim firstChoice As Object
        Set firstChoice = choices(1)
        Set GetAssistantMessage = firstChoice("message")
    Else
        Set GetAssistantMessage = Nothing
    End If
    Exit Function
ErrHandler:
    Set GetAssistantMessage = Nothing
End Function
