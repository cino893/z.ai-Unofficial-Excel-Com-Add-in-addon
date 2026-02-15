' ============================================================================
' build.vbs - Build script for Z.AI Excel Add-in
' Creates .xlam file by importing VBA modules into Excel via COM automation
'
' Usage: cscript build.vbs
'        or double-click to run
' ============================================================================
Option Explicit

Dim fso, shell, excel, wb, vbProj
Dim scriptDir, outputPath
Dim modulesImported

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
outputPath = fso.BuildPath(scriptDir, "ZaiExcelAddin.xlam")

WScript.Echo "============================================="
WScript.Echo " Z.AI Excel Add-in Builder"
WScript.Echo "============================================="
WScript.Echo ""
WScript.Echo "Katalog zrodlowy: " & scriptDir
WScript.Echo "Plik wyjsciowy:   " & outputPath
WScript.Echo ""

' Check if output exists
If fso.FileExists(outputPath) Then
    WScript.Echo "Usuwam istniejacy plik: " & outputPath
    fso.DeleteFile outputPath, True
End If

' Start Excel
WScript.Echo "Uruchamiam Excel..."
On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "BLAD: Nie mozna uruchomic Excel. Upewnij sie ze Excel jest zainstalowany."
    WScript.Quit 1
End If
On Error GoTo 0

excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False

' Check VBA trust settings
On Error Resume Next
Set wb = excel.Workbooks.Add
If Err.Number <> 0 Then
    WScript.Echo "BLAD: Nie mozna utworzyc skoroszytu."
    excel.Quit
    WScript.Quit 1
End If

Set vbProj = wb.VBProject
If Err.Number <> 0 Then
    WScript.Echo ""
    WScript.Echo "BLAD: Brak dostepu do projektu VBA!"
    WScript.Echo ""
    WScript.Echo "Musisz wlaczyc dostep do modelu obiektow VBA:"
    WScript.Echo "1. Otworz Excel"
    WScript.Echo "2. Plik > Opcje > Centrum zaufania > Ustawienia Centrum zaufania"
    WScript.Echo "3. Ustawienia makr > Zaznacz 'Ufaj dostepowi do modelu obiektow projektu VBA'"
    WScript.Echo "4. Kliknij OK i uruchom ten skrypt ponownie"
    WScript.Echo ""
    wb.Close False
    excel.Quit
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "Dostep do VBA: OK"
WScript.Echo ""

modulesImported = 0

' Import modules
ImportModule "modJSON.bas"
ImportModule "modDebug.bas"
ImportModule "modAuth.bas"
ImportModule "modZaiAPI.bas"
ImportModule "modExcelSkills.bas"
ImportModule "modConversation.bas"
ImportModule "modRibbon.bas"

' Create the chat UserForm programmatically
WScript.Echo "Tworzenie formularza czatu..."
CreateChatForm vbProj

' Add auto-open code to ThisWorkbook
WScript.Echo "Dodawanie kodu auto-open..."
AddThisWorkbookCode vbProj

WScript.Echo ""
WScript.Echo "Zaimportowano modulow: " & modulesImported
WScript.Echo ""

' Save as .xlam
WScript.Echo "Zapisywanie jako .xlam..."
On Error Resume Next
' 55 = xlAddIn (.xlam)
wb.SaveAs outputPath, 55
If Err.Number <> 0 Then
    WScript.Echo "BLAD przy zapisywaniu: " & Err.Description
    ' Try alternate format number
    Err.Clear
    wb.SaveAs outputPath, 52 ' xlOpenXMLWorkbookMacroEnabled
    If Err.Number <> 0 Then
        WScript.Echo "BLAD: Nie udalo sie zapisac pliku. " & Err.Description
        wb.Close False
        excel.Quit
        WScript.Quit 1
    End If
End If
On Error GoTo 0

wb.Close False
excel.Quit

WScript.Echo ""
WScript.Echo "============================================="
WScript.Echo " SUKCES! Dodatek zostal utworzony:"
WScript.Echo " " & outputPath
WScript.Echo "============================================="
WScript.Echo ""
WScript.Echo "Aby zainstalowac dodatek:"
WScript.Echo "1. Otworz Excel"
WScript.Echo "2. Plik > Opcje > Dodatki"
WScript.Echo "3. Na dole: Zarzadzaj > Dodatki programu Excel > Przejdz"
WScript.Echo "4. Kliknij Przegladaj i wskaż: " & outputPath
WScript.Echo "5. Zaznacz 'ZaiExcelAddin' i kliknij OK"
WScript.Echo ""
WScript.Echo "Menu Z.AI pojawi sie na pasku menu Excel."

' ============================================================================
' Helper: Import a .bas module
' ============================================================================
Sub ImportModule(fileName)
    Dim filePath
    filePath = fso.BuildPath(scriptDir, fileName)
    
    If Not fso.FileExists(filePath) Then
        WScript.Echo "  UWAGA: Pominięto brakujacy plik: " & fileName
        Exit Sub
    End If
    
    On Error Resume Next
    vbProj.VBComponents.Import filePath
    If Err.Number <> 0 Then
        WScript.Echo "  BLAD importu " & fileName & ": " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Zaimportowano: " & fileName
        modulesImported = modulesImported + 1
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' Helper: Create Chat UserForm programmatically
' ============================================================================
Sub CreateChatForm(proj)
    On Error Resume Next
    
    Dim frm, ctrl
    Set frm = proj.VBComponents.Add(3) ' vbext_ct_MSForm = 3
    
    If Err.Number <> 0 Then
        WScript.Echo "  BLAD: Nie mozna dodac UserForm: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    frm.Name = "frmZaiChat"
    If Err.Number <> 0 Then Err.Clear
    
    ' Set form properties via Designer
    With frm.Designer
        .Caption = "Z.AI - Asystent Excel"
        .Width = 520
        .Height = 620
    End With
    If Err.Number <> 0 Then Err.Clear
    
    ' Chat history TextBox
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1", "txtChat")
    With ctrl
        .Left = 10
        .Top = 10
        .Width = 490
        .Height = 470
        .MultiLine = True
        .ScrollBars = 2 ' fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
        .Font.Name = "Consolas"
        .Font.Size = 10
    End With
    
    ' Status label
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1", "lblStatus")
    With ctrl
        .Left = 10
        .Top = 490
        .Width = 490
        .Height = 18
        .Caption = "Gotowy"
        .Font.Size = 8
        .ForeColor = &H808080
    End With
    
    ' Input TextBox
    Set ctrl = frm.Designer.Controls.Add("Forms.TextBox.1", "txtInput")
    With ctrl
        .Left = 10
        .Top = 512
        .Width = 400
        .Height = 50
        .MultiLine = True
        .WordWrap = True
        .EnterKeyBehavior = False
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
    
    ' Send button
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1", "btnSend")
    With ctrl
        .Left = 418
        .Top = 512
        .Width = 82
        .Height = 50
        .Caption = "Wyslij"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' New conversation button
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1", "btnNew")
    With ctrl
        .Left = 10
        .Top = 570
        .Width = 120
        .Height = 25
        .Caption = "Nowa rozmowa"
        .Font.Size = 8
    End With
    
    ' Clear button
    Set ctrl = frm.Designer.Controls.Add("Forms.CommandButton.1", "btnClear")
    With ctrl
        .Left = 140
        .Top = 570
        .Width = 100
        .Height = 25
        .Caption = "Wyczysc"
        .Font.Size = 8
    End With
    
    ' Add form code
    Dim code
    code = ""
    code = code & "Option Explicit" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "#If VBA7 Then" & vbCrLf
    code = code & "Private Declare PtrSafe Function GetWindowLong Lib ""user32"" Alias ""GetWindowLongA"" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long" & vbCrLf
    code = code & "Private Declare PtrSafe Function SetWindowLong Lib ""user32"" Alias ""SetWindowLongA"" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long" & vbCrLf
    code = code & "Private Declare PtrSafe Function FindWindow Lib ""user32"" Alias ""FindWindowA"" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr" & vbCrLf
    code = code & "Private Declare PtrSafe Function DrawMenuBar Lib ""user32"" (ByVal hWnd As LongPtr) As Long" & vbCrLf
    code = code & "#Else" & vbCrLf
    code = code & "Private Declare Function GetWindowLong Lib ""user32"" Alias ""GetWindowLongA"" (ByVal hWnd As Long, ByVal nIndex As Long) As Long" & vbCrLf
    code = code & "Private Declare Function SetWindowLong Lib ""user32"" Alias ""SetWindowLongA"" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long" & vbCrLf
    code = code & "Private Declare Function FindWindow Lib ""user32"" Alias ""FindWindowA"" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long" & vbCrLf
    code = code & "Private Declare Function DrawMenuBar Lib ""user32"" (ByVal hWnd As Long) As Long" & vbCrLf
    code = code & "#End If" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Const GWL_STYLE As Long = -16" & vbCrLf
    code = code & "Private Const WS_THICKFRAME As Long = &H40000" & vbCrLf
    code = code & "Private Const WS_MAXIMIZEBOX As Long = &H10000" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private m_conversationStarted As Boolean" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    Me.Width = 520" & vbCrLf
    code = code & "    Me.Height = 620" & vbCrLf
    code = code & "    Me.Caption = ""Z.AI - Asystent Excel""" & vbCrLf
    code = code & "    m_conversationStarted = False" & vbCrLf
    code = code & "    AppendChat ""Z.AI"", ""Witaj! Jestem asystentem AI zintegrowanym z Excel."" & vbCrLf & _" & vbCrLf
    code = code & "        ""Moge pomoc Ci edytowac arkusz - po prostu opisz co chcesz zrobic."" & vbCrLf & vbCrLf & _" & vbCrLf
    code = code & "        ""Przyklady polecen:"" & vbCrLf & _" & vbCrLf
    code = code & "        ""  - Przeczytaj dane z kolumny A"" & vbCrLf & _" & vbCrLf
    code = code & "        ""  - Dodaj formule SUM do komorki B10"" & vbCrLf & _" & vbCrLf
    code = code & "        ""  - Sformatuj naglowki na pogrubione"" & vbCrLf & _" & vbCrLf
    code = code & "        ""  - Posortuj dane wedlug kolumny C malejaco"" & vbCrLf & _" & vbCrLf
    code = code & "        ""  - Stworz wykres kolumnowy z danych A1:B5""" & vbCrLf
    code = code & "    modDebug.LogInfo ""Chat form opened""" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub UserForm_Activate()" & vbCrLf
    code = code & "    MakeResizable" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub MakeResizable()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim hWnd As LongPtr" & vbCrLf
    code = code & "    hWnd = FindWindow(""ThunderDFrame"", Me.Caption)" & vbCrLf
    code = code & "    If hWnd = 0 Then Exit Sub" & vbCrLf
    code = code & "    Dim style As Long" & vbCrLf
    code = code & "    style = GetWindowLong(hWnd, GWL_STYLE)" & vbCrLf
    code = code & "    style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX" & vbCrLf
    code = code & "    SetWindowLong hWnd, GWL_STYLE, style" & vbCrLf
    code = code & "    DrawMenuBar hWnd" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub UserForm_Resize()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    Dim w As Single, h As Single" & vbCrLf
    code = code & "    w = Me.InsideWidth" & vbCrLf
    code = code & "    h = Me.InsideHeight" & vbCrLf
    code = code & "    If w < 300 Or h < 200 Then Exit Sub" & vbCrLf
    code = code & "    Dim inputH As Single: inputH = 50" & vbCrLf
    code = code & "    Dim statusH As Single: statusH = 18" & vbCrLf
    code = code & "    Dim btnRowH As Single: btnRowH = 25" & vbCrLf
    code = code & "    Dim pad As Single: pad = 10" & vbCrLf
    code = code & "    Dim sendW As Single: sendW = 82" & vbCrLf
    code = code & "    txtChat.Left = pad" & vbCrLf
    code = code & "    txtChat.Top = pad" & vbCrLf
    code = code & "    txtChat.Width = w - 2 * pad" & vbCrLf
    code = code & "    txtChat.Height = h - pad - statusH - 4 - inputH - 4 - btnRowH - pad - pad" & vbCrLf
    code = code & "    lblStatus.Left = pad" & vbCrLf
    code = code & "    lblStatus.Top = txtChat.Top + txtChat.Height + 4" & vbCrLf
    code = code & "    lblStatus.Width = w - 2 * pad" & vbCrLf
    code = code & "    txtInput.Left = pad" & vbCrLf
    code = code & "    txtInput.Top = lblStatus.Top + statusH + 4" & vbCrLf
    code = code & "    txtInput.Width = w - 2 * pad - sendW - 8" & vbCrLf
    code = code & "    txtInput.Height = inputH" & vbCrLf
    code = code & "    btnSend.Left = txtInput.Left + txtInput.Width + 8" & vbCrLf
    code = code & "    btnSend.Top = txtInput.Top" & vbCrLf
    code = code & "    btnSend.Width = sendW" & vbCrLf
    code = code & "    btnSend.Height = inputH" & vbCrLf
    code = code & "    btnNew.Left = pad" & vbCrLf
    code = code & "    btnNew.Top = txtInput.Top + inputH + 4" & vbCrLf
    code = code & "    btnClear.Left = btnNew.Left + btnNew.Width + 10" & vbCrLf
    code = code & "    btnClear.Top = btnNew.Top" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub btnSend_Click()" & vbCrLf
    code = code & "    SendMessage" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub btnNew_Click()" & vbCrLf
    code = code & "    modConversation.ResetConversation" & vbCrLf
    code = code & "    txtChat.Text = """"" & vbCrLf
    code = code & "    m_conversationStarted = False" & vbCrLf
    code = code & "    AppendChat ""Z.AI"", ""Nowa rozmowa rozpoczeta. Jak moge pomoc?""" & vbCrLf
    code = code & "    lblStatus.Caption = ""Gotowy""" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub btnClear_Click()" & vbCrLf
    code = code & "    txtChat.Text = """"" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf
    code = code & "    If KeyCode = 13 And Shift = 0 Then" & vbCrLf
    code = code & "        KeyCode = 0" & vbCrLf
    code = code & "        SendMessage" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub SendMessage()" & vbCrLf
    code = code & "    Dim userText As String" & vbCrLf
    code = code & "    userText = Trim(txtInput.Text)" & vbCrLf
    code = code & "    If userText = """" Then Exit Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    If Not m_conversationStarted Then" & vbCrLf
    code = code & "        modConversation.InitConversation" & vbCrLf
    code = code & "        m_conversationStarted = True" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    AppendChat ""Ty"", userText" & vbCrLf
    code = code & "    txtInput.Text = """"" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    txtInput.Enabled = False" & vbCrLf
    code = code & "    btnSend.Enabled = False" & vbCrLf
    code = code & "    lblStatus.Caption = ""Przetwarzanie...""" & vbCrLf
    code = code & "    DoEvents" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    Dim response As String" & vbCrLf
    code = code & "    response = modConversation.SendUserMessage(userText)" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    AppendChat ""Z.AI"", response" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    txtInput.Enabled = True" & vbCrLf
    code = code & "    btnSend.Enabled = True" & vbCrLf
    code = code & "    lblStatus.Caption = ""Gotowy ("" & modConversation.GetMessageCount() & "" wiadomosci)""" & vbCrLf
    code = code & "    txtInput.SetFocus" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub AppendChat(ByVal sender As String, ByVal message As String)" & vbCrLf
    code = code & "    Dim ts As String" & vbCrLf
    code = code & "    ts = Format(Now, ""hh:nn"")" & vbCrLf
    code = code & "    Dim newText As String" & vbCrLf
    code = code & "    If txtChat.Text <> """" Then newText = txtChat.Text & vbCrLf & vbCrLf" & vbCrLf
    code = code & "    newText = newText & ""["" & ts & ""] "" & sender & "":"" & vbCrLf & message" & vbCrLf
    code = code & "    txtChat.Text = newText" & vbCrLf
    code = code & "    txtChat.SelStart = Len(txtChat.Text)" & vbCrLf
    code = code & "    txtChat.SelLength = 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub UserForm_Terminate()" & vbCrLf
    code = code & "    modDebug.LogInfo ""Chat form closed""" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    ' Insert code into the form module
    frm.CodeModule.DeleteLines 1, frm.CodeModule.CountOfLines
    frm.CodeModule.AddFromString code
    
    If Err.Number <> 0 Then
        WScript.Echo "  UWAGA: Blad przy tworzeniu formularza: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Utworzono formularz: frmZaiChat"
        modulesImported = modulesImported + 1
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' Helper: Add ThisWorkbook auto-open code
' ============================================================================
Sub AddThisWorkbookCode(proj)
    On Error Resume Next
    
    Dim twb, comp
    Dim i
    ' Find ThisWorkbook component (Type=100 = Document, name varies by locale)
    For i = 1 To proj.VBComponents.Count
        Set comp = proj.VBComponents(i)
        If comp.Type = 100 Then
            Dim codeName
            codeName = LCase(comp.Name)
            If InStr(codeName, "workbook") > 0 Or _
               InStr(codeName, "skoroszyt") > 0 Then
                Set twb = comp
                Exit For
            End If
        End If
    Next
    
    ' Fallback: first Type=100 that is not a sheet
    If twb Is Nothing Then
        For i = 1 To proj.VBComponents.Count
            Set comp = proj.VBComponents(i)
            If comp.Type = 100 Then
                If Not (LCase(comp.Name) Like "ark*" Or _
                        LCase(comp.Name) Like "sheet*") Then
                    Set twb = comp
                    Exit For
                End If
            End If
        Next
    End If
    
    If twb Is Nothing Then
        WScript.Echo "  BLAD: Nie znaleziono ThisWorkbook w projekcie VBA"
        Exit Sub
    End If
    
    WScript.Echo "  Znaleziono ThisWorkbook jako: " & twb.Name
    
    Dim code
    code = ""
    code = code & "Private Sub Workbook_Open()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    modDebug.InitDebug" & vbCrLf
    code = code & "    modDebug.LogInfo ""Z.AI Add-in loaded""" & vbCrLf
    code = code & "    modRibbon.CreateZaiMenu" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub Workbook_AddinInstall()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    modDebug.InitDebug" & vbCrLf
    code = code & "    modDebug.LogInfo ""Z.AI Add-in installed""" & vbCrLf
    code = code & "    modRibbon.CreateZaiMenu" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub Workbook_AddinUninstall()" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    modRibbon.RemoveZaiMenu" & vbCrLf
    code = code & "    modDebug.LogInfo ""Z.AI Add-in uninstalled""" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    modRibbon.RemoveZaiMenu" & vbCrLf
    code = code & "    modDebug.LogInfo ""Z.AI Add-in unloaded""" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    twb.CodeModule.DeleteLines 1, twb.CodeModule.CountOfLines
    twb.CodeModule.AddFromString code
    
    If Err.Number <> 0 Then
        WScript.Echo "  UWAGA: Blad przy dodawaniu kodu ThisWorkbook: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "  Dodano kod ThisWorkbook (auto-open/close)"
    End If
    
    On Error GoTo 0
End Sub
