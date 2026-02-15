Attribute VB_Name = "modAuth"
'==============================================================================
' modAuth - Authentication Module
' Manages z.ai API key storage via Windows Registry
'==============================================================================
Option Explicit

Private Const REG_KEY As String = "HKCU\Software\ZaiExcelAddin\"
Private Const REG_VAL_APIKEY As String = "ApiKey"
Private Const REG_VAL_MODEL As String = "Model"

Private m_cachedApiKey As String

' --- Save API key to Registry ---
Public Sub SaveApiKey(ByVal apiKey As String)
    On Error GoTo ErrHandler
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_KEY & REG_VAL_APIKEY, apiKey, "REG_SZ"
    m_cachedApiKey = apiKey
    LogInfo "API key saved to registry"
    Exit Sub
ErrHandler:
    LogError "Failed to save API key: " & Err.Description
End Sub

' --- Load API key from Registry ---
Public Function LoadApiKey() As String
    If m_cachedApiKey <> "" Then
        LoadApiKey = m_cachedApiKey
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    m_cachedApiKey = wsh.RegRead(REG_KEY & REG_VAL_APIKEY)
    LoadApiKey = m_cachedApiKey
    LogDebug "API key loaded from registry"
    Exit Function
ErrHandler:
    LoadApiKey = ""
    LogDebug "No API key found in registry"
End Function

' --- Clear API key ---
Public Sub ClearApiKey()
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegDelete REG_KEY & REG_VAL_APIKEY
    m_cachedApiKey = ""
    LogInfo "API key cleared"
End Sub

' --- Check if logged in ---
Public Function IsLoggedIn() As Boolean
    IsLoggedIn = (LoadApiKey() <> "")
End Function

' --- Save preferred model ---
Public Sub SaveModel(ByVal modelName As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_KEY & REG_VAL_MODEL, modelName, "REG_SZ"
End Sub

' --- Load preferred model ---
Public Function LoadModel() As String
    On Error GoTo ErrHandler
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    LoadModel = wsh.RegRead(REG_KEY & REG_VAL_MODEL)
    Exit Function
ErrHandler:
    LoadModel = "glm-4-plus"
End Function

' --- Validate API key by making a test call ---
Public Function ValidateApiKey(ByVal apiKey As String) As Boolean
    LogInfo "Validating API key..."
    
    On Error GoTo ErrHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = ZAI_API_BASE & "/chat/completions"
    
    Dim body As String
    body = "{""model"":""glm-4-plus"",""messages"":[{""role"":""user"",""content"":""Hi""}],""max_tokens"":5}"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send body
    
    LogDebug "Validation response: HTTP " & http.Status
    
    If http.Status = 200 Then
        ValidateApiKey = True
        LogInfo "API key is valid"
    ElseIf http.Status = 401 Then
        ValidateApiKey = False
        LogWarn "API key is invalid (401 Unauthorized)"
    Else
        ' Might still be valid - some models may not be available
        ValidateApiKey = (http.Status <> 401 And http.Status <> 403)
        LogWarn "Validation returned HTTP " & http.Status & ": " & Left(http.responseText, 500)
    End If
    
    Exit Function
ErrHandler:
    LogError "Validation failed: " & Err.Description
    ValidateApiKey = False
End Function

' --- Show login dialog ---
Public Sub ShowLogin()
    Dim apiKey As String
    Dim currentKey As String
    currentKey = LoadApiKey()
    
    Dim prompt As String
    prompt = "Podaj klucz API z.ai:" & vbCrLf & vbCrLf & _
             "Klucz mozesz uzyskac na: https://open.z.ai/" & vbCrLf & _
             "(Sekcja: API Keys)"
    
    If currentKey <> "" Then
        prompt = prompt & vbCrLf & vbCrLf & "Aktualny klucz: " & Left(currentKey, 8) & "..." & Right(currentKey, 4)
    End If
    
    apiKey = InputBox(prompt, "Z.AI - Logowanie")
    
    If apiKey = "" Then
        If currentKey = "" Then
            MsgBox "Logowanie anulowane. Musisz podac klucz API aby korzystac z Z.AI.", vbExclamation, "Z.AI"
        End If
        Exit Sub
    End If
    
    ' Clean up the key
    apiKey = Trim(apiKey)
    
    ' Validate
    Application.StatusBar = "Z.AI: Weryfikacja klucza API..."
    Dim isValid As Boolean
    isValid = ValidateApiKey(apiKey)
    Application.StatusBar = False
    
    If isValid Then
        SaveApiKey apiKey
        MsgBox "Zalogowano pomyslnie!" & vbCrLf & "Klucz API zostal zapisany.", vbInformation, "Z.AI"
    Else
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Nie udalo sie zweryfikowac klucza API." & vbCrLf & _
                      "Czy chcesz go mimo to zapisac?", vbQuestion + vbYesNo, "Z.AI")
        If resp = vbYes Then
            SaveApiKey apiKey
        End If
    End If
End Sub

' --- Show logout confirmation ---
Public Sub ShowLogout()
    If Not IsLoggedIn() Then
        MsgBox "Nie jestes zalogowany.", vbInformation, "Z.AI"
        Exit Sub
    End If
    
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Czy na pewno chcesz sie wylogowac?" & vbCrLf & _
                  "Klucz API zostanie usuniety.", vbQuestion + vbYesNo, "Z.AI"
    )
    If resp = vbYes Then
        ClearApiKey
        MsgBox "Wylogowano.", vbInformation, "Z.AI"
    End If
End Sub
