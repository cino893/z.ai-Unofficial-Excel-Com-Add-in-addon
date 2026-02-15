Attribute VB_Name = "modConversation"
'==============================================================================
' modConversation - Conversation Management with Tool-Calling Loop
' Manages multi-turn chat with z.ai agent, executing tool calls automatically
'==============================================================================
Option Explicit

Private Const MAX_TOOL_ROUNDS As Long = 15
Private Const MAX_SAME_TOOL_REPEATS As Long = 2
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
    GetSystemPrompt = T("system.prompt")
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
    
    ' Tool-calling loop with repetition detection
    Dim round As Long
    Dim finalResponse As String
    finalResponse = ""
    Dim lastToolSig As String
    lastToolSig = ""
    Dim sameToolCount As Long
    sameToolCount = 0
    
    For round = 1 To MAX_TOOL_ROUNDS
        LogInfo "=== Conversation round " & round & " ==="
        
        ' Send to API
        Application.StatusBar = TFormat("conv.status_round", round)
        DoEvents
        
        Dim messagesJson As String
        messagesJson = BuildMessagesJson()
        
        Dim toolsJson As String
        toolsJson = GetToolDefinitions()
        
        Dim response As Object
        Set response = SendChatCompletion(messagesJson, toolsJson)
        
        If response Is Nothing Or Not response("success") Then
            finalResponse = T("conv.api_error") & DictGet(response, "error", T("conv.no_response"))
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
            finalResponse = T("conv.error") & T("conv.no_assistant")
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
            
            ' Detect repetitive tool calls
            Dim currentSig As String
            currentSig = BuildToolCallSignature(toolCalls)
            If currentSig = lastToolSig And currentSig <> "" Then
                sameToolCount = sameToolCount + 1
                If sameToolCount >= MAX_SAME_TOOL_REPEATS Then
                    LogWarn "Repetitive tool call detected (" & sameToolCount & "x): " & currentSig
                    finalResponse = T("conv.loop_detected")
                    GoTo Cleanup
                End If
            Else
                sameToolCount = 0
                lastToolSig = currentSig
            End If
            
            Application.StatusBar = TFormat("conv.status_exec", toolCalls.Count)
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
    finalResponse = TFormat("conv.max_rounds", MAX_TOOL_ROUNDS)
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

' --- Build signature string for tool calls to detect repetition ---
Private Function BuildToolCallSignature(ByVal toolCalls As Collection) As String
    On Error GoTo ErrHandler
    Dim sig As String
    sig = ""
    Dim tc As Long
    For tc = 1 To toolCalls.Count
        Dim toolCall As Object
        Set toolCall = toolCalls(tc)
        Dim funcObj As Object
        Set funcObj = toolCall("function")
        If sig <> "" Then sig = sig & "|"
        sig = sig & CStr(funcObj("name")) & ":" & CStr(funcObj("arguments"))
    Next tc
    BuildToolCallSignature = sig
    Exit Function
ErrHandler:
    BuildToolCallSignature = ""
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
