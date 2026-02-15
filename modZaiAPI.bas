Attribute VB_Name = "modZaiAPI"
'==============================================================================
' modZaiAPI - z.ai API Communication Module
' Handles HTTP requests to z.ai Chat Completions API
'==============================================================================
Option Explicit

Public Const ZAI_API_BASE As String = "https://api.z.ai/api/paas/v4"
Public Const ZAI_DEFAULT_MODEL As String = "glm-4-plus"
Public Const ZAI_MAX_TOKENS As Long = 4096
Public Const ZAI_TEMPERATURE As Double = 0.7

' --- Send Chat Completion Request ---
Public Function SendChatCompletion( _
    ByVal messagesJson As String, _
    Optional ByVal toolsJson As String = "", _
    Optional ByVal model As String = "" _
) As Object
    
    If model = "" Then model = LoadModel()
    
    Dim apiKey As String
    apiKey = LoadApiKey()
    If apiKey = "" Then
        LogError "SendChatCompletion: No API key"
        Set SendChatCompletion = Nothing
        Exit Function
    End If
    
    ' Build request body
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
    
    ' Send request
    Dim url As String
    url = ZAI_API_BASE & "/chat/completions"
    
    Dim response As Object
    Set response = HttpPost(url, body, apiKey)
    Set SendChatCompletion = response
End Function

' --- HTTP POST ---
Public Function HttpPost(ByVal url As String, ByVal body As String, ByVal apiKey As String) As Object
    On Error GoTo ErrHandler
    
    LogApiRequest "POST", url, body
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "Accept-Language", "pl-PL,pl;q=0.9,en-US;q=0.8,en;q=0.7"
    http.send body
    
    LogApiResponse http.Status, http.responseText
    
    ' Parse response
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("status") = http.Status
    result("raw") = http.responseText
    
    If http.Status = 200 Then
        Dim parsed As Object
        Set parsed = JsonParse(http.responseText)
        If Not parsed Is Nothing Then
            Set result("data") = parsed
        End If
        result("success") = True
    Else
        result("success") = False
        result("error") = "HTTP " & http.Status & ": " & Left(http.responseText, 500)
        LogError "API error: " & result("error")
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

' --- Extract assistant message content from response ---
Public Function GetResponseContent(ByVal response As Object) As String
    On Error GoTo ErrHandler
    
    If response Is Nothing Then
        GetResponseContent = ""
        Exit Function
    End If
    
    If Not response("success") Then
        GetResponseContent = "[Blad]: " & DictGet(response, "error", "Nieznany blad")
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
    Set firstChoice = choices(1) ' Collection is 1-based
    
    Dim msg As Object
    Set msg = firstChoice("message")
    
    If msg.Exists("content") Then
        If Not IsNull(msg("content")) Then
            GetResponseContent = CStr(msg("content"))
        End If
    End If
    
    Exit Function
ErrHandler:
    LogErrorDetails "GetResponseContent", Err.Number, Err.Description
    GetResponseContent = ""
End Function

' --- Extract tool calls from response ---
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

' --- Check if response has tool calls ---
Public Function HasToolCalls(ByVal response As Object) As Boolean
    On Error GoTo ErrHandler
    
    Dim calls As Collection
    Set calls = GetToolCalls(response)
    HasToolCalls = (calls.Count > 0)
    Exit Function
    
ErrHandler:
    HasToolCalls = False
End Function

' --- Get finish reason ---
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

' --- Get full assistant message dict from response ---
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
