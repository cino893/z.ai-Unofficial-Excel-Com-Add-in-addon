Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - Excel Menu Bar Integration
' Creates Z.AI menu in Excel's toolbar
'==============================================================================
Option Explicit

Private Const MENU_NAME As String = "Z.AI"

' --- Create menu on add-in load ---
Public Sub CreateZaiMenu()
    On Error Resume Next
    ' Remove existing menu first
    RemoveZaiMenu
    On Error GoTo ErrHandler
    
    ' Initialize i18n
    InitI18n
    
    Dim menuBar As CommandBar
    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    
    Dim zaiMenu As CommandBarPopup
    Set zaiMenu = menuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    zaiMenu.Caption = MENU_NAME
    
    ' --- Chat (main feature) ---
    Dim btnChat As CommandBarButton
    Set btnChat = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnChat.Caption = T("menu.chat")
    btnChat.FaceId = 581
    btnChat.OnAction = "ShowChatDialog"
    btnChat.BeginGroup = False
    
    ' --- Quick Command ---
    Dim btnQuick As CommandBarButton
    Set btnQuick = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnQuick.Caption = T("menu.quick")
    btnQuick.FaceId = 487
    btnQuick.OnAction = "ShowQuickCommand"
    
    ' --- Separator + Auth ---
    Dim btnLogin As CommandBarButton
    Set btnLogin = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnLogin.Caption = T("menu.login")
    btnLogin.FaceId = 2144
    btnLogin.OnAction = "ShowLogin"
    btnLogin.BeginGroup = True
    
    Dim btnLogout As CommandBarButton
    Set btnLogout = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnLogout.Caption = T("menu.logout")
    btnLogout.FaceId = 358
    btnLogout.OnAction = "ShowLogout"
    
    ' --- Separator + Settings ---
    Dim btnModel As CommandBarButton
    Set btnModel = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnModel.Caption = T("menu.model")
    btnModel.FaceId = 548
    btnModel.OnAction = "ShowModelSelector"
    btnModel.BeginGroup = True
    
    ' --- Language ---
    Dim btnLang As CommandBarButton
    Set btnLang = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnLang.Caption = T("menu.language")
    btnLang.FaceId = 224
    btnLang.OnAction = "ShowLanguageSelector"
    
    ' --- Separator + Debug ---
    Dim btnViewLog As CommandBarButton
    Set btnViewLog = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnViewLog.Caption = T("menu.viewlog")
    btnViewLog.FaceId = 547
    btnViewLog.OnAction = "ViewLog"
    btnViewLog.BeginGroup = True
    
    Dim btnClearLog As CommandBarButton
    Set btnClearLog = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnClearLog.Caption = T("menu.clearlog")
    btnClearLog.FaceId = 472
    btnClearLog.OnAction = "ClearLog"
    
    ' --- About ---
    Dim btnAbout As CommandBarButton
    Set btnAbout = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnAbout.Caption = T("menu.about")
    btnAbout.FaceId = 487
    btnAbout.OnAction = "ShowAbout"
    btnAbout.BeginGroup = True
    
    LogInfo "Z.AI menu created"
    Exit Sub
    
ErrHandler:
    LogErrorDetails "CreateZaiMenu", Err.Number, Err.Description
End Sub

' --- Remove menu ---
Public Sub RemoveZaiMenu()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(MENU_NAME).Delete
    On Error GoTo 0
End Sub

' ======================== MENU CALLBACKS ========================

' --- Show Chat Dialog ---
Public Sub ShowChatDialog()
    On Error GoTo ErrHandler
    
    If Not IsLoggedIn() Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox(T("auth.need_login") & vbCrLf & _
                      T("auth.want_login"), vbQuestion + vbYesNo, "Z.AI")
        If resp = vbYes Then
            ShowLogin
            If Not IsLoggedIn() Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' Show the chat form
    frmZaiChat.Show vbModeless
    
    Exit Sub
ErrHandler:
    LogErrorDetails "ShowChatDialog", Err.Number, Err.Description
    MsgBox T("error.generic") & Err.Description, vbCritical, "Z.AI"
End Sub

' --- Show Quick Command ---
Public Sub ShowQuickCommand()
    On Error GoTo ErrHandler
    
    If Not IsLoggedIn() Then
        ShowLogin
        If Not IsLoggedIn() Then Exit Sub
    End If
    
    Dim command As String
    command = InputBox(T("quick.prompt"), T("quick.title"))
    
    If command = "" Then Exit Sub
    
    Application.StatusBar = T("quick.status")
    DoEvents
    
    Dim result As String
    result = QuickCommand(command)
    
    Application.StatusBar = False
    
    MsgBox result, vbInformation, T("quick.result_title")
    
    Exit Sub
ErrHandler:
    Application.StatusBar = False
    LogErrorDetails "ShowQuickCommand", Err.Number, Err.Description
    MsgBox T("error.generic") & Err.Description, vbCritical, "Z.AI"
End Sub

' --- Show Model Selector ---
Public Sub ShowModelSelector()
    Dim currentModel As String
    currentModel = LoadModel()
    
    Dim model As String
    model = InputBox( _
        T("model.prompt") & vbCrLf & vbCrLf & _
        T("model.current") & currentModel, _
        T("model.title"), _
        currentModel)
    
    If model = "" Then Exit Sub
    
    SaveModel Trim(model)
    MsgBox T("model.changed") & Trim(model), vbInformation, "Z.AI"
End Sub

' --- Show Language Selector ---
Public Sub ShowLanguageSelector()
    Dim current As String
    current = GetLanguage()
    
    Dim choice As String
    choice = InputBox( _
        "Select language / Wybierz jezyk:" & vbCrLf & vbCrLf & _
        "  pl - Polski" & vbCrLf & _
        "  en - English" & vbCrLf & vbCrLf & _
        "Current / Aktualny: " & current, _
        T("lang.title"), _
        current)
    
    If choice = "" Then Exit Sub
    
    choice = LCase(Trim(choice))
    If choice = "pl" Or choice = "en" Then
        SetLanguage choice
        ' Recreate menu with new language
        CreateZaiMenu
        MsgBox T("lang.changed"), vbInformation, T("lang.title")
    End If
End Sub

' --- Show About ---
Public Sub ShowAbout()
    MsgBox T("about.text"), vbInformation, T("about.title")
End Sub
