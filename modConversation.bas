Attribute VB_Name = "modConversation"
'==============================================================================
' modConversation - Conversation Management with Tool-Calling Loop
' Manages multi-turn chat with z.ai agent, executing tool calls automatically
'==============================================================================
Option Explicit

Private Const MAX_TOOL_ROUNDS As Long = 15
Private m_messages As Collection
Private m_isProcessing As Boolean

' --- Initialize conversation ---
Public Sub InitConversation()
    Set m_messages = New Collection
    m_isProcessing = False
    
    ' Add system prompt
    Dim sysMsg As Object
    Set sysMsg = CreateObject("Scripting.Dictionary")
    sysMsg("role") = "system"
    sysMsg("content") = GetSystemPrompt()
    m_messages.Add sysMsg
    
    LogInfo "Conversation initialized with system prompt"
End Sub

' --- Get system prompt ---
Public Function GetSystemPrompt() As String
    GetSystemPrompt = _
        "Jestes inteligentnym asystentem AI zintegrowanym z Microsoft Excel. " & _
        "Pomagasz uzytkownikowi edytowac i analizowac dane w arkuszu kalkulacyjnym. " & _
        "Masz dostep do narzedzi (tools) ktore pozwalaja Ci czytac i zapisywac komorki, " & _
        "formatowac dane, wstawiac formuly, sortowac, tworzyc wykresy i wiele wiecej." & vbLf & vbLf & _
        "ZASADY:" & vbLf & _
        "1. Zawsze NAJPIERW uzyj get_sheet_info lub get_workbook_info aby poznac kontekst danych." & vbLf & _
        "2. Przed modyfikacja danych, przeczytaj odpowiedni zakres aby zrozumiec strukture." & vbLf & _
        "3. Po wykonaniu zmian, potwierdzaj co zrobiles." & vbLf & _
        "4. Uzywaj polskich nazw w komunikatach do uzytkownika." & vbLf & _
        "5. Formuly Excel pisz w skladni angielskiej (SUM, AVERAGE, IF, VLOOKUP itp.)." & vbLf & _
        "6. Kolory podawaj jako RGB long: Red=255, Green=65280, Blue=16711680, Yellow=65535, " & _
        "LightGray=12632256, White=16777215, Orange=33023." & vbLf & _
        "7. Jezeli uzytkownik nie sprecyzuje arkusza, uzyj aktywnego arkusza." & vbLf & _
        "8. Badz zwiezly ale informatywny w odpowiedziach."
End Function

' --- Build messages JSON ---
Public Function BuildMessagesJson() As String
    Dim result As String
    result = "["
    
    Dim i As Long
    For i = 1 To m_messages.Count
        If i > 1 Then result = result & ","
        
        Dim msg As Object
        Set msg = m_messages(i)
        
        result = result & "{""role"":""" & msg("role") & """"
        
        If msg.Exists("content") Then
            If Not IsNull(msg("content")) And msg("content") <> "" Then
                result = result & ",""content"":""" & EscapeJsonString(CStr(msg("content"))) & """"
            Else
                result = result & ",""content"":null"
            End If
        End If
        
        If msg.Exists("tool_calls_json") Then
            result = result & ",""tool_calls"":" & msg("tool_calls_json")
        End If
        
        If msg.Exists("tool_call_id") Then
            result = result & ",""tool_call_id"":""" & msg("tool_call_id") & """"
        End If
        
        result = result & "}"
    Next i
    
    result = result & "]"
    BuildMessagesJson = result
End Function

' --- Send user message and run full tool-calling loop ---
Public Function SendUserMessage(ByVal userMessage As String) As String
    On Error GoTo ErrHandler
    
    If m_messages Is Nothing Then InitConversation
    
    LogInfo "User message: " & userMessage
    
    ' Add user message
    Dim userMsg As Object
    Set userMsg = CreateObject("Scripting.Dictionary")
    userMsg("role") = "user"
    userMsg("content") = userMessage
    m_messages.Add userMsg
    
    m_isProcessing = True
    
    ' Tool-calling loop
    Dim round As Long
    Dim finalResponse As String
    finalResponse = ""
    
    For round = 1 To MAX_TOOL_ROUNDS
        LogInfo "=== Conversation round " & round & " ==="
        
        ' Send to API
        Application.StatusBar = "Z.AI: Przetwarzanie... (runda " & round & ")"
        DoEvents
        
        Dim messagesJson As String
        messagesJson = BuildMessagesJson()
        
        Dim toolsJson As String
        toolsJson = GetToolDefinitions()
        
        Dim response As Object
        Set response = SendChatCompletion(messagesJson, toolsJson)
        
        If response Is Nothing Or Not response("success") Then
            finalResponse = "[Blad API]: " & DictGet(response, "error", "Brak odpowiedzi z serwera")
            LogError "API call failed in round " & round
            GoTo Cleanup
        End If
        
        Dim finishReason As String
        finishReason = GetFinishReason(response)
        LogDebug "Finish reason: " & finishReason
        
        ' Get assistant message and add to history
        Dim assistantMsg As Object
        Set assistantMsg = GetAssistantMessage(response)
        
        If assistantMsg Is Nothing Then
            finalResponse = "[Blad]: Brak odpowiedzi asystenta"
            GoTo Cleanup
        End If
        
        ' Check for tool calls
        If HasToolCalls(response) Then
            ' Add assistant message with tool calls to history
            Dim asstDict As Object
            Set asstDict = CreateObject("Scripting.Dictionary")
            asstDict("role") = "assistant"
            
            If assistantMsg.Exists("content") Then
                If Not IsNull(assistantMsg("content")) Then
                    asstDict("content") = CStr(assistantMsg("content"))
                Else
                    asstDict("content") = ""
                End If
            Else
                asstDict("content") = ""
            End If
            
            ' Serialize tool_calls back to JSON for the message history
            Dim toolCallsJson As String
            toolCallsJson = SerializeToolCalls(assistantMsg("tool_calls"))
            asstDict("tool_calls_json") = toolCallsJson
            m_messages.Add asstDict
            
            ' Execute tool calls
            Dim toolCalls As Collection
            Set toolCalls = GetToolCalls(response)
            
            Application.StatusBar = "Z.AI: Wykonywanie " & toolCalls.Count & " operacji na Excelu..."
            DoEvents
            
            Dim tc As Long
            For tc = 1 To toolCalls.Count
                Dim toolCall As Object
                Set toolCall = toolCalls(tc)
                
                Dim funcObj As Object
                Set funcObj = toolCall("function")
                
                Dim tcId As String
                tcId = CStr(toolCall("id"))
                
                Dim funcName As String
                funcName = CStr(funcObj("name"))
                
                Dim funcArgs As String
                funcArgs = CStr(funcObj("arguments"))
                
                LogInfo "Executing tool: " & funcName & " (id: " & tcId & ")"
                
                ' Execute the skill
                Dim toolResult As String
                toolResult = ExecuteToolCall(funcName, funcArgs)
                
                ' Add tool result to messages
                Dim toolMsg As Object
                Set toolMsg = CreateObject("Scripting.Dictionary")
                toolMsg("role") = "tool"
                toolMsg("content") = toolResult
                toolMsg("tool_call_id") = tcId
                m_messages.Add toolMsg
            Next tc
            
            ' Continue the loop - agent needs to process tool results
        Else
            ' No tool calls - we have the final response
            finalResponse = GetResponseContent(response)
            
            ' Add assistant message to history
            Dim finalAsstDict As Object
            Set finalAsstDict = CreateObject("Scripting.Dictionary")
            finalAsstDict("role") = "assistant"
            finalAsstDict("content") = finalResponse
            m_messages.Add finalAsstDict
            
            GoTo Cleanup
        End If
    Next round
    
    ' Max rounds reached
    finalResponse = "[Uwaga]: Osiagnieto maksymalna liczbe rund (" & MAX_TOOL_ROUNDS & "). Ostatnia odpowiedz moze byc niekompletna."
    LogWarn "Max tool-calling rounds reached"
    
Cleanup:
    m_isProcessing = False
    Application.StatusBar = False
    SendUserMessage = finalResponse
    Exit Function
    
ErrHandler:
    m_isProcessing = False
    Application.StatusBar = False
    LogErrorDetails "SendUserMessage", Err.Number, Err.Description
    SendUserMessage = "[Blad]: " & Err.Description
End Function

' --- Serialize tool_calls collection back to JSON ---
Private Function SerializeToolCalls(ByVal toolCalls As Object) As String
    On Error GoTo ErrHandler
    
    Dim result As String
    result = "["
    
    Dim i As Long
    For i = 1 To toolCalls.Count
        If i > 1 Then result = result & ","
        
        Dim tc As Object
        Set tc = toolCalls(i)
        
        Dim funcObj As Object
        Set funcObj = tc("function")
        
        result = result & "{""id"":""" & EscapeJsonString(CStr(tc("id"))) & """," & _
                 """type"":""function""," & _
                 """function"":{""name"":""" & EscapeJsonString(CStr(funcObj("name"))) & """," & _
                 """arguments"":""" & EscapeJsonString(CStr(funcObj("arguments"))) & """}}"
    Next i
    
    result = result & "]"
    SerializeToolCalls = result
    Exit Function
    
ErrHandler:
    LogErrorDetails "SerializeToolCalls", Err.Number, Err.Description
    SerializeToolCalls = "[]"
End Function

' --- Reset conversation ---
Public Sub ResetConversation()
    InitConversation
    LogInfo "Conversation reset"
End Sub

' --- Check if processing ---
Public Function IsProcessing() As Boolean
    IsProcessing = m_isProcessing
End Function

' --- Get conversation message count ---
Public Function GetMessageCount() As Long
    If m_messages Is Nothing Then
        GetMessageCount = 0
    Else
        GetMessageCount = m_messages.Count
    End If
End Function

' --- Quick command (one-shot, no chat history) ---
Public Function QuickCommand(ByVal command As String) As String
    InitConversation
    QuickCommand = SendUserMessage(command)
End Function
