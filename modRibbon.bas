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
    
    Dim menuBar As CommandBar
    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    
    Dim zaiMenu As CommandBarPopup
    Set zaiMenu = menuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    zaiMenu.Caption = MENU_NAME
    
    ' --- Chat (main feature) ---
    Dim btnChat As CommandBarButton
    Set btnChat = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnChat.Caption = "&Asystent AI (Chat)"
    btnChat.FaceId = 581
    btnChat.OnAction = "ShowChatDialog"
    btnChat.BeginGroup = False
    
    ' --- Quick Command ---
    Dim btnQuick As CommandBarButton
    Set btnQuick = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnQuick.Caption = "&Szybkie polecenie"
    btnQuick.FaceId = 487
    btnQuick.OnAction = "ShowQuickCommand"
    
    ' --- Separator + Auth ---
    Dim btnLogin As CommandBarButton
    Set btnLogin = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnLogin.Caption = "&Zaloguj (Klucz API)"
    btnLogin.FaceId = 2144
    btnLogin.OnAction = "ShowLogin"
    btnLogin.BeginGroup = True
    
    Dim btnLogout As CommandBarButton
    Set btnLogout = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnLogout.Caption = "&Wyloguj"
    btnLogout.FaceId = 358
    btnLogout.OnAction = "ShowLogout"
    
    ' --- Separator + Settings ---
    Dim btnModel As CommandBarButton
    Set btnModel = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnModel.Caption = "&Wybierz model"
    btnModel.FaceId = 548
    btnModel.OnAction = "ShowModelSelector"
    btnModel.BeginGroup = True
    
    ' --- Separator + Debug ---
    Dim btnViewLog As CommandBarButton
    Set btnViewLog = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnViewLog.Caption = "&Pokaz log debugowania"
    btnViewLog.FaceId = 547
    btnViewLog.OnAction = "ViewLog"
    btnViewLog.BeginGroup = True
    
    Dim btnClearLog As CommandBarButton
    Set btnClearLog = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnClearLog.Caption = "&Wyczysc log"
    btnClearLog.FaceId = 472
    btnClearLog.OnAction = "ClearLog"
    
    ' --- About ---
    Dim btnAbout As CommandBarButton
    Set btnAbout = zaiMenu.Controls.Add(Type:=msoControlButton)
    btnAbout.Caption = "&O dodatku Z.AI"
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
        resp = MsgBox("Musisz najpierw sie zalogowac (podac klucz API)." & vbCrLf & _
                      "Czy chcesz to zrobic teraz?", vbQuestion + vbYesNo, "Z.AI")
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
    MsgBox "Blad: " & Err.Description, vbCritical, "Z.AI"
End Sub

' --- Show Quick Command ---
Public Sub ShowQuickCommand()
    On Error GoTo ErrHandler
    
    If Not IsLoggedIn() Then
        ShowLogin
        If Not IsLoggedIn() Then Exit Sub
    End If
    
    Dim command As String
    command = InputBox( _
        "Wpisz polecenie dla asystenta AI:" & vbCrLf & vbCrLf & _
        "Przyklady:" & vbCrLf & _
        "  - Podsumuj dane w kolumnie A" & vbCrLf & _
        "  - Dodaj formule SUM do B10" & vbCrLf & _
        "  - Sformatuj naglowki na pogrubione" & vbCrLf & _
        "  - Stworz wykres z danych A1:B10", _
        "Z.AI - Szybkie polecenie")
    
    If command = "" Then Exit Sub
    
    Application.StatusBar = "Z.AI: Przetwarzanie polecenia..."
    DoEvents
    
    Dim result As String
    result = QuickCommand(command)
    
    Application.StatusBar = False
    
    MsgBox result, vbInformation, "Z.AI - Odpowiedz"
    
    Exit Sub
ErrHandler:
    Application.StatusBar = False
    LogErrorDetails "ShowQuickCommand", Err.Number, Err.Description
    MsgBox "Blad: " & Err.Description, vbCritical, "Z.AI"
End Sub

' --- Show Model Selector ---
Public Sub ShowModelSelector()
    Dim currentModel As String
    currentModel = LoadModel()
    
    Dim model As String
    model = InputBox( _
        "Wybierz model z.ai:" & vbCrLf & vbCrLf & _
        "Dostepne modele:" & vbCrLf & _
        "  glm-4-plus  (domyslny, szybki)" & vbCrLf & _
        "  glm-4-long  (dlugi kontekst)" & vbCrLf & _
        "  glm-4       (standardowy)" & vbCrLf & _
        "  glm-3-turbo (najszybszy)" & vbCrLf & vbCrLf & _
        "Aktualny: " & currentModel, _
        "Z.AI - Wybor modelu", _
        currentModel)
    
    If model = "" Then Exit Sub
    
    SaveModel Trim(model)
    MsgBox "Model zmieniony na: " & Trim(model), vbInformation, "Z.AI"
End Sub

' --- Show About ---
Public Sub ShowAbout()
    MsgBox _
        "Z.AI Excel Add-in" & vbCrLf & _
        "Wersja: 1.0.0" & vbCrLf & vbCrLf & _
        "Asystent AI zintegrowany z Microsoft Excel." & vbCrLf & _
        "Wykorzystuje platforme z.ai (Zhipu AI) do" & vbCrLf & _
        "inteligentnej edycji arkuszy kalkulacyjnych." & vbCrLf & vbCrLf & _
        "Mozliwosci:" & vbCrLf & _
        "  - Czytanie i zapisywanie komorek" & vbCrLf & _
        "  - Formatowanie danych" & vbCrLf & _
        "  - Wstawianie formul" & vbCrLf & _
        "  - Sortowanie danych" & vbCrLf & _
        "  - Tworzenie wykresow" & vbCrLf & _
        "  - Zarzadzanie arkuszami" & vbCrLf & vbCrLf & _
        "Strona: https://z.ai" & vbCrLf & _
        "Dokumentacja API: https://docs.z.ai", _
        vbInformation, "Z.AI - O dodatku"
End Sub
