Attribute VB_Name = "modDebug"
'==============================================================================
' modDebug - Debug Logging Module
' Logs to file + Immediate window with timestamps and levels
'==============================================================================
Option Explicit

Public Enum LogLevel
    LOG_DEBUG = 0
    LOG_INFO = 1
    LOG_WARN = 2
    LOG_ERROR = 3
End Enum

Private m_logLevel As LogLevel
Private m_logFilePath As String
Private m_initialized As Boolean

' --- Initialize logging ---
Public Sub InitDebug(Optional ByVal minLevel As LogLevel = LOG_DEBUG)
    m_logLevel = minLevel
    m_logFilePath = GetLogFilePath()
    m_initialized = True
    LogInfo "=== Z.AI Excel Add-in Debug Log Started ==="
    LogInfo "Log file: " & m_logFilePath
    LogInfo "Excel version: " & Application.Version
    LogInfo "OS: " & Application.OperatingSystem
End Sub

' --- Get log file path ---
Public Function GetLogFilePath() As String
    Dim folder As String
    folder = Environ("APPDATA") & "\ZaiExcelAddin"
    
    ' Create folder if needed
    If Dir(folder, vbDirectory) = "" Then
        MkDir folder
    End If
    
    GetLogFilePath = folder & "\zai_debug_" & Format(Date, "yyyy-mm-dd") & ".log"
End Function

' --- Core logging function ---
Public Sub DebugLog(ByVal message As String, Optional ByVal level As LogLevel = LOG_INFO)
    If Not m_initialized Then InitDebug
    If level < m_logLevel Then Exit Sub
    
    Dim levelStr As String
    Select Case level
        Case LOG_DEBUG: levelStr = "DEBUG"
        Case LOG_INFO: levelStr = "INFO "
        Case LOG_WARN: levelStr = "WARN "
        Case LOG_ERROR: levelStr = "ERROR"
    End Select
    
    Dim logLine As String
    logLine = Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & levelStr & "] " & message
    
    ' Print to Immediate window
    Debug.Print logLine
    
    ' Write to file
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    Open m_logFilePath For Append As #fileNum
    Print #fileNum, logLine
    Close #fileNum
    On Error GoTo 0
End Sub

' --- Convenience methods ---
Public Sub LogDebug(ByVal message As String)
    DebugLog message, LOG_DEBUG
End Sub

Public Sub LogInfo(ByVal message As String)
    DebugLog message, LOG_INFO
End Sub

Public Sub LogWarn(ByVal message As String)
    DebugLog message, LOG_WARN
End Sub

Public Sub LogError(ByVal message As String)
    DebugLog message, LOG_ERROR
End Sub

' --- Log API request ---
Public Sub LogApiRequest(ByVal method As String, ByVal url As String, ByVal body As String)
    LogDebug ">>> API REQUEST: " & method & " " & url
    ' Truncate body for readability
    If Len(body) > 2000 Then
        LogDebug ">>> BODY (truncated): " & Left(body, 2000) & "..."
    Else
        LogDebug ">>> BODY: " & body
    End If
End Sub

' --- Log API response ---
Public Sub LogApiResponse(ByVal statusCode As Long, ByVal responseText As String)
    LogDebug "<<< API RESPONSE: HTTP " & statusCode
    If Len(responseText) > 3000 Then
        LogDebug "<<< BODY (truncated): " & Left(responseText, 3000) & "..."
    Else
        LogDebug "<<< BODY: " & responseText
    End If
End Sub

' --- Log tool call ---
Public Sub LogToolCall(ByVal toolName As String, ByVal args As String, ByVal result As String)
    LogInfo "TOOL CALL: " & toolName
    LogDebug "  Args: " & args
    If Len(result) > 1000 Then
        LogDebug "  Result (truncated): " & Left(result, 1000) & "..."
    Else
        LogDebug "  Result: " & result
    End If
End Sub

' --- Clear log file ---
Public Sub ClearLog()
    If Not m_initialized Then InitDebug
    On Error Resume Next
    Kill m_logFilePath
    On Error GoTo 0
    LogInfo "Log cleared"
End Sub

' --- Open log file in Notepad ---
Public Sub ViewLog()
    If Not m_initialized Then InitDebug
    If Dir(m_logFilePath) <> "" Then
        Shell "notepad.exe """ & m_logFilePath & """", vbNormalFocus
    Else
        MsgBox T("debug.no_log") & vbCrLf & T("debug.path") & m_logFilePath, vbInformation, "Z.AI Debug"
    End If
End Sub

' --- Log error with full details ---
Public Sub LogErrorDetails(ByVal source As String, ByVal errNum As Long, ByVal errDesc As String)
    LogError "EXCEPTION in " & source & ": [" & errNum & "] " & errDesc
End Sub

' --- Get log folder path ---
Public Function GetLogFolder() As String
    GetLogFolder = Environ("APPDATA") & "\ZaiExcelAddin"
End Function
