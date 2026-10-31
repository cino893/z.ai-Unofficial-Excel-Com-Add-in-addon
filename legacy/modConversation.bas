Attribute VB_Name = "modConversation"
'==============================================================================
' modConversation - Conversation Management with Tool-Calling Loop
' Mirrors src/Services/ConversationService.cs
'==============================================================================
Option Explicit

Public Const MAX_TOOL_ROUNDS As Long = 45
Private Const MAX_SAME_TOOL_REPEATS As Long = 2
Private Const SOFT_LIMIT As Long = 80
Private Const XL_CALC_MANUAL As Long = -4135

Public Const SR_NONE As Long = 0
Public Const SR_COMPLETED As Long = 1
Public Const SR_MAX_ROUNDS As Long = 2
Public Const SR_LOOP_DETECTED As Long = 3
Public Const SR_CANCELLED As Long = 4
Public Const SR_ERROR As Long = 5

Private m_messages As Collection
Private m_isProcessing As Boolean
Private m_cancelRequested As Boolean
Private m_lastStopReason As Long
Private m_lastAssistantResponse As String

' --- Initialize conversation ---
Public Sub InitConversation()
    Set m_messages = New Collection
    m_isProcessing = False
    m_cancelRequested = False
    m_lastStopReason = SR_NONE
    m_lastAssistantResponse = ""
    
    Dim sysMsg As Object
    Set sysMsg = CreateObject("Scripting.Dictionary")
    sysMsg("role") = "system"
    sysMsg("content") = GetSystemPrompt()
    m_messages.Add sysMsg
    
    LogInfo "Conversation initialized with system prompt"
End Sub

Public Function GetSystemPrompt() As String
    Dim repl As Object
    Set repl = CreateObject("Scripting.Dictionary")
    repl("MaxToolRounds") = CStr(MAX_TOOL_ROUNDS)
    GetSystemPrompt = T("system.prompt", repl)
End Function

Public Sub CancelConversation()
    m_cancelRequested = True
    LogInfo "Cancellation requested"
End Sub

Public Function IsCancelRequested() As Boolean
    IsCancelRequested = m_cancelRequested
End Function

Public Function LastStopReason() As Long
    LastStopReason = m_lastStopReason
End Function

Public Function CanContinue() As Boolean
    CanContinue = (m_lastStopReason = SR_MAX_ROUNDS Or m_lastStopReason = SR_CANCELLED)
End Function

Public Function LastAssistantResponse() As String
    LastAssistantResponse = m_lastAssistantResponse
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
            If Not IsNull(msg("content")) And CStr(msg("content")) <> "" Then
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

Public Function SendUserMessage(ByVal userMessage As String) As String
    On Error GoTo ErrHandler
    
    If m_messages Is Nothing Then InitConversation
    
    m_isProcessing = True
    m_cancelRequested = False
    m_lastStopReason = SR_NONE
    
    TrimOldTurns
    
    Dim userMsg As Object
    Set userMsg = CreateObject("Scripting.Dictionary")
    userMsg("role") = "user"
    userMsg("content") = userMessage
    m_messages.Add userMsg
    
    LogInfo "User message added, history size: " & m_messages.Count
    SendUserMessage = RunToolLoop()
    m_isProcessing = False
    Exit Function
    
ErrHandler:
    m_isProcessing = False
    m_lastStopReason = SR_ERROR
    LogErrorDetails "SendUserMessage", Err.Number, Err.Description
    SendUserMessage = "Error: " & Err.Description
End Function

Public Function ContinueConversation() As String
    On Error GoTo ErrHandler
    
    If m_messages Is Nothing Then InitConversation
    
    m_isProcessing = True
    m_cancelRequested = False
    
    TrimOldTurns
    
    Dim useLastSummary As Boolean
    useLastSummary = (m_lastStopReason = SR_MAX_ROUNDS And Trim$(m_lastAssistantResponse) <> "")
    m_lastStopReason = SR_NONE
    
    Dim continueMsg As String
    If useLastSummary Then
        continueMsg = m_lastAssistantResponse
    Else
        continueMsg = T("conv.continue_prompt")
        Dim recentContext As String
        recentContext = BuildRecentContext()
        If recentContext <> "" Then
            continueMsg = continueMsg & vbLf & vbLf & "Last completed actions:" & vbLf & recentContext
        End If
    End If
    
    Dim userMsg As Object
    Set userMsg = CreateObject("Scripting.Dictionary")
    userMsg("role") = "user"
    userMsg("content") = continueMsg
    m_messages.Add userMsg
    
    LogInfo "Continue requested, resuming tool loop"
    ContinueConversation = RunToolLoop()
    m_isProcessing = False
    Exit Function
    
ErrHandler:
    m_isProcessing = False
    m_lastStopReason = SR_ERROR
    LogErrorDetails "ContinueConversation", Err.Number, Err.Description
    ContinueConversation = "Error: " & Err.Description
End Function

Private Function RunToolLoop() As String
    Dim oldSU As Boolean, oldEE As Boolean, oldCalc As Long
    oldSU = True
    oldEE = True
    oldCalc = xlCalculationAutomatic
    On Error Resume Next
    oldSU = Application.ScreenUpdating
    oldEE = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = XL_CALC_MANUAL
    On Error GoTo 0
    
    Dim result As String
    Dim loopErr As Long
    Dim loopDesc As String
    On Error Resume Next
    result = RunToolLoopInner()
    loopErr = Err.Number
    loopDesc = Err.Description
    Err.Clear
    
    On Error Resume Next
    Application.ScreenUpdating = oldSU
    On Error Resume Next
    Application.EnableEvents = oldEE
    On Error Resume Next
    Application.Calculation = oldCalc
    On Error GoTo 0
    
    Application.StatusBar = False
    If loopErr <> 0 Then
        m_lastStopReason = SR_ERROR
        result = "Error: " & loopDesc
    End If
    RunToolLoop = result
End Function

Private Function RunToolLoopInner() As String
    Dim previousSignature As String, previousToolNames As String
    previousSignature = ""
    previousToolNames = ""
    Dim repeatCount As Long, nameRepeatCount As Long
    Dim roundInfoIndex As Long, finalPromptIndex As Long
    roundInfoIndex = 0
    finalPromptIndex = 0
    
    Dim round As Long
    For round = 1 To MAX_TOOL_ROUNDS
        DoEvents
        If m_cancelRequested Then
            LogInfo "Cancelled by user"
            m_lastStopReason = SR_CANCELLED
            RemoveTransientMessages roundInfoIndex, finalPromptIndex
            RunToolLoopInner = T("conv.cancelled")
            Exit Function
        End If
        
        LogInfo "Tool-calling loop round " & round & "/" & MAX_TOOL_ROUNDS
        Application.StatusBar = TFormat("conv.status_round", round)
        
        Dim repl As Object
        Set repl = CreateObject("Scripting.Dictionary")
        repl("CurrentRound") = CStr(round)
        repl("MaxToolRounds") = CStr(MAX_TOOL_ROUNDS)
        
        Dim roundMsg As Object
        Set roundMsg = CreateObject("Scripting.Dictionary")
        roundMsg("role") = "system"
        roundMsg("content") = T("conv.round_info", repl)
        
        If roundInfoIndex >= 1 And roundInfoIndex <= m_messages.Count Then
            m_messages.Remove roundInfoIndex
            If finalPromptIndex > roundInfoIndex Then finalPromptIndex = finalPromptIndex - 1
        End If
        m_messages.Add roundMsg
        roundInfoIndex = m_messages.Count
        
        If round = MAX_TOOL_ROUNDS Then
            Dim finalRepl As Object
            Set finalRepl = CreateObject("Scripting.Dictionary")
            finalRepl("MaxToolRounds") = CStr(MAX_TOOL_ROUNDS)
            Dim finalMsg As Object
            Set finalMsg = CreateObject("Scripting.Dictionary")
            finalMsg("role") = "user"
            finalMsg("content") = T("conv.final_round_prompt", finalRepl)
            m_messages.Add finalMsg
            finalPromptIndex = m_messages.Count
        End If
        
        Dim messagesJson As String, toolsJson As String
        messagesJson = BuildMessagesJson()
        toolsJson = GetToolDefinitions()
        
        Dim response As Object
        Set response = SendChatCompletion(messagesJson, toolsJson)
        
        If response Is Nothing Or Not DictGet(response, "success", False) Then
            If m_cancelRequested Or response Is Nothing Then
                m_lastStopReason = SR_CANCELLED
                LogInfo "Cancelled during API call"
                RemoveTransientMessages roundInfoIndex, finalPromptIndex
                RunToolLoopInner = T("conv.cancelled")
                Exit Function
            End If
            If DictGet(response, "cancelled", False) Then
                m_lastStopReason = SR_CANCELLED
                RemoveTransientMessages roundInfoIndex, finalPromptIndex
                RunToolLoopInner = T("conv.cancelled")
                Exit Function
            End If
            m_lastStopReason = SR_ERROR
            Dim apiErr As String
            apiErr = DictGet(response, "error", T("conv.api_error"))
            LogError "API call failed: " & apiErr
            RemoveTransientMessages roundInfoIndex, finalPromptIndex
            RunToolLoopInner = CStr(apiErr)
            Exit Function
        End If
        
        Dim finishReason As String
        finishReason = GetFinishReason(response)
        LogInfo "Finish reason: " & finishReason
        
        If HasToolCalls(response) Then
            Dim assistantMsg As Object
            Set assistantMsg = GetAssistantMessage(response)
            
            Dim asstDict As Object
            Set asstDict = CreateObject("Scripting.Dictionary")
            asstDict("role") = "assistant"
            If Not assistantMsg Is Nothing Then
                If assistantMsg.Exists("content") Then
                    If Not IsNull(assistantMsg("content")) Then
                        asstDict("content") = CStr(assistantMsg("content"))
                    Else
                        asstDict("content") = ""
                    End If
                Else
                    asstDict("content") = ""
                End If
            Else
                asstDict("content") = ""
            End If
            
            Dim toolCalls As Collection
            Set toolCalls = GetToolCalls(response)
            asstDict("tool_calls_json") = SerializeToolCalls(toolCalls)
            m_messages.Add asstDict
            
            Dim signature As String, toolNames As String
            signature = BuildToolCallSignature(toolCalls)
            toolNames = BuildToolNamesSignature(toolCalls)
            
            Dim isLoop As Boolean
            isLoop = False
            
            If signature = previousSignature And signature <> "" Then
                repeatCount = repeatCount + 1
                LogWarn "Same tool signature repeated (" & repeatCount & "/" & MAX_SAME_TOOL_REPEATS & ")"
                If repeatCount >= MAX_SAME_TOOL_REPEATS Then isLoop = True
            Else
                repeatCount = 0
            End If
            previousSignature = signature
            
            If toolNames = previousToolNames And toolNames <> "" Then
                nameRepeatCount = nameRepeatCount + 1
                If nameRepeatCount >= MAX_SAME_TOOL_REPEATS + 1 Then
                    LogWarn "Same tool names repeated " & nameRepeatCount & " times"
                    isLoop = True
                End If
            Else
                nameRepeatCount = 0
            End If
            previousToolNames = toolNames
            
            If isLoop Then
                LogWarn "Tool loop detected, breaking out"
                Dim dtc As Long
                For dtc = 1 To toolCalls.Count
                    Dim dummy As Object
                    Set dummy = CreateObject("Scripting.Dictionary")
                    dummy("role") = "tool"
                    dummy("content") = "{""error"":""Loop detected, execution stopped""}"
                    dummy("tool_call_id") = CStr(toolCalls(dtc)("id"))
                    m_messages.Add dummy
                Next dtc
                m_lastStopReason = SR_LOOP_DETECTED
                RemoveTransientMessages roundInfoIndex, finalPromptIndex
                RunToolLoopInner = T("conv.loop_detected")
                Exit Function
            End If
            
            Application.StatusBar = TFormat("conv.status_exec", toolCalls.Count)
            DoEvents
            
            Dim tc As Long
            For tc = 1 To toolCalls.Count
                Dim toolCall As Object
                Set toolCall = toolCalls(tc)
                Dim funcObj As Object
                Set funcObj = toolCall("function")
                Dim tcId As String, funcName As String, funcArgs As String
                tcId = CStr(toolCall("id"))
                funcName = CStr(funcObj("name"))
                funcArgs = CStr(funcObj("arguments"))
                
                Dim toolResult As String
                If m_cancelRequested Then
                    toolResult = "{""error"":""Cancelled by user""}"
                Else
                    LogInfo "Executing tool: " & funcName
                    toolResult = ExecuteToolCall(funcName, funcArgs)
                    DoEvents
                End If
                
                Dim toolMsg As Object
                Set toolMsg = CreateObject("Scripting.Dictionary")
                toolMsg("role") = "tool"
                toolMsg("content") = toolResult
                toolMsg("tool_call_id") = tcId
                m_messages.Add toolMsg
            Next tc
        Else
            Dim content As String
            content = GetResponseContent(response)
            If Trim$(content) = "" Then
                m_lastStopReason = SR_ERROR
                LogError "Assistant returned empty response"
                RemoveTransientMessages roundInfoIndex, finalPromptIndex
                RunToolLoopInner = T("conv.generic_failure")
                Exit Function
            End If
            RemoveTransientMessages roundInfoIndex, finalPromptIndex
            Dim finalAsst As Object
            Set finalAsst = CreateObject("Scripting.Dictionary")
            finalAsst("role") = "assistant"
            finalAsst("content") = content
            m_messages.Add finalAsst
            m_lastAssistantResponse = content
            LogInfo "Assistant response received (" & Len(content) & " chars)"
            m_lastStopReason = SR_COMPLETED
            RunToolLoopInner = content
            Exit Function
        End If
    Next round
    
    RemoveTransientMessages roundInfoIndex, finalPromptIndex
    LogWarn "Max tool rounds (" & MAX_TOOL_ROUNDS & ") reached"
    m_lastStopReason = SR_MAX_ROUNDS
    RunToolLoopInner = TFormat("conv.max_rounds", MAX_TOOL_ROUNDS)
End Function

Private Sub RemoveTransientMessages(ByRef roundInfoIndex As Long, ByRef finalPromptIndex As Long)
    If finalPromptIndex >= 1 And finalPromptIndex <= m_messages.Count Then
        m_messages.Remove finalPromptIndex
        If roundInfoIndex > finalPromptIndex Then roundInfoIndex = roundInfoIndex - 1
        finalPromptIndex = 0
    End If
    If roundInfoIndex >= 1 And roundInfoIndex <= m_messages.Count Then
        m_messages.Remove roundInfoIndex
        roundInfoIndex = 0
    End If
End Sub

Private Function BuildRecentContext() As String
    If m_messages Is Nothing Then
        BuildRecentContext = ""
        Exit Function
    End If
    
    Dim lines As New Collection
    Dim i As Long
    For i = m_messages.Count To 1 Step -1
        If lines.Count >= 10 Then Exit For
        Dim msg As Object
        Set msg = m_messages(i)
        Dim role As String
        role = CStr(msg("role"))
        If role = "tool" Then
            Dim content As String
            content = CStr(msg("content"))
            If Len(content) > 100 Then content = Left$(content, 100) & "..."
            lines.Add "  → result: " & content
        ElseIf role = "assistant" And msg.Exists("tool_calls_json") Then
            lines.Add "  • called: (tools)"
        ElseIf role = "user" Then
            Exit For
        End If
    Next i
    
    Dim result As String
    result = ""
    For i = lines.Count To 1 Step -1
        If result <> "" Then result = result & vbLf
        result = result & lines(i)
    Next i
    BuildRecentContext = result
End Function

Private Sub TrimOldTurns()
    If m_messages Is Nothing Then Exit Sub
    If m_messages.Count <= SOFT_LIMIT Then Exit Sub
    
    Dim userIndices As New Collection
    Dim i As Long
    For i = 2 To m_messages.Count
        If CStr(m_messages(i)("role")) = "user" Then userIndices.Add i
    Next i
    
    Do While m_messages.Count > SOFT_LIMIT And userIndices.Count >= 2
        Dim fromIdx As Long, toIdx As Long, count As Long
        fromIdx = userIndices(1)
        toIdx = userIndices(2)
        count = toIdx - fromIdx
        Dim j As Long
        For j = 1 To count
            m_messages.Remove fromIdx
        Next j
        userIndices.Remove 1
        Dim k As Long
        Dim tmp As New Collection
        For k = 1 To userIndices.Count
            tmp.Add userIndices(k) - count
        Next k
        Set userIndices = tmp
        LogInfo "Trimmed " & count & " old messages, history now: " & m_messages.Count
    Loop
End Sub

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

Private Function BuildToolNamesSignature(ByVal toolCalls As Collection) As String
    On Error GoTo ErrHandler
    Dim names() As String
    Dim n As Long
    n = toolCalls.Count
    If n = 0 Then
        BuildToolNamesSignature = ""
        Exit Function
    End If
    ReDim names(1 To n)
    Dim tc As Long
    For tc = 1 To n
        names(tc) = CStr(toolCalls(tc)("function")("name"))
    Next tc
    ' Simple sort so name-level detection is order-independent (matches COM OrderBy)
    Dim i As Long, j As Long, tmp As String
    For i = 1 To n - 1
        For j = i + 1 To n
            If names(j) < names(i) Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next j
    Next i
    Dim sig As String
    sig = names(1)
    For i = 2 To n
        sig = sig & "|" & names(i)
    Next i
    BuildToolNamesSignature = sig
    Exit Function
ErrHandler:
    BuildToolNamesSignature = ""
End Function

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

Public Sub ResetConversation()
    InitConversation
    LogInfo "Conversation reset"
End Sub

Public Function IsProcessing() As Boolean
    IsProcessing = m_isProcessing
End Function

Public Function GetMessageCount() As Long
    If m_messages Is Nothing Then
        GetMessageCount = 0
    Else
        GetMessageCount = m_messages.Count
    End If
End Function

Public Function QuickCommand(ByVal command As String) As String
    InitConversation
    QuickCommand = SendUserMessage(command)
End Function
