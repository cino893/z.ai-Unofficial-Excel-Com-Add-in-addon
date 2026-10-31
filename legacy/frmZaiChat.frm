VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZaiChat
   Caption         =   "Z.AI - Excel Assistant"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "frmZaiChat.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frmZaiChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmZaiChat - Chat Interface for Z.AI Agent
' Provides conversational UI for interacting with the AI agent
'==============================================================================
Option Explicit

Private m_conversationStarted As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = T("chat.title")
    Me.Width = 520
    Me.Height = 640
    
    CreateControls
    UserForm_Resize
    
    m_conversationStarted = False
    AppendChat "Z.AI", T("chat.welcome")
    LogInfo "Chat form opened"
End Sub

Private Sub CreateControls()
    On Error Resume Next
    
    Dim txtChat As MSForms.TextBox
    Set txtChat = Me.Controls.Add("Forms.TextBox.1", "txtChat")
    With txtChat
        .Left = 10
        .Top = 10
        .Width = 490
        .Height = 450
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
        .Font.Name = "Consolas"
        .Font.Size = 10
        .BackColor = &H00FFFFFF
    End With
    
    Dim lblStatus As MSForms.Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Left = 10
        .Top = 466
        .Width = 490
        .Height = 18
        .Caption = T("chat.ready")
        .Font.Size = 8
        .ForeColor = &H00808080
    End With
    
    Dim txtInput As MSForms.TextBox
    Set txtInput = Me.Controls.Add("Forms.TextBox.1", "txtInput")
    With txtInput
        .Left = 10
        .Top = 488
        .Width = 400
        .Height = 50
        .MultiLine = True
        .WordWrap = True
        .EnterKeyBehavior = False
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
    
    Dim btnSend As MSForms.CommandButton
    Set btnSend = Me.Controls.Add("Forms.CommandButton.1", "btnSend")
    With btnSend
        .Left = 418
        .Top = 488
        .Width = 82
        .Height = 50
        .Caption = T("chat.send")
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    Dim btnStop As MSForms.CommandButton
    Set btnStop = Me.Controls.Add("Forms.CommandButton.1", "btnStop")
    With btnStop
        .Left = 418
        .Top = 488
        .Width = 82
        .Height = 50
        .Caption = T("chat.stop")
        If .Caption = "chat.stop" Then .Caption = "Stop"
        .Font.Size = 10
        .Font.Bold = True
        .Visible = False
        .ForeColor = &H000000C0
    End With
    
    Dim btnNew As MSForms.CommandButton
    Set btnNew = Me.Controls.Add("Forms.CommandButton.1", "btnNew")
    With btnNew
        .Left = 10
        .Top = 546
        .Width = 110
        .Height = 25
        .Caption = T("chat.new")
        .Font.Size = 8
    End With
    
    Dim btnClear As MSForms.CommandButton
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear")
    With btnClear
        .Left = 128
        .Top = 546
        .Width = 90
        .Height = 25
        .Caption = T("chat.clear")
        .Font.Size = 8
    End With
    
    Dim btnContinue As MSForms.CommandButton
    Set btnContinue = Me.Controls.Add("Forms.CommandButton.1", "btnContinue")
    With btnContinue
        .Left = 226
        .Top = 546
        .Width = 110
        .Height = 25
        .Caption = T("chat.continue")
        If .Caption = "chat.continue" Then .Caption = "Continue"
        .Font.Size = 8
        .Visible = False
    End With
    
    On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    Dim w As Single, h As Single
    w = Me.InsideWidth
    h = Me.InsideHeight
    If w < 240 Or h < 240 Then Exit Sub
    
    Dim pad As Single, inputH As Single, statusH As Single, btnRowH As Single, sendW As Single
    pad = 10
    inputH = 50
    statusH = 18
    btnRowH = 25
    sendW = 82
    
    Me.Controls("txtChat").Left = pad
    Me.Controls("txtChat").Top = pad
    Me.Controls("txtChat").Width = w - 2 * pad
    Me.Controls("txtChat").Height = h - pad - statusH - 4 - inputH - 4 - btnRowH - pad - pad
    
    Me.Controls("lblStatus").Left = pad
    Me.Controls("lblStatus").Top = Me.Controls("txtChat").Top + Me.Controls("txtChat").Height + 4
    Me.Controls("lblStatus").Width = w - 2 * pad
    
    Me.Controls("txtInput").Left = pad
    Me.Controls("txtInput").Top = Me.Controls("lblStatus").Top + statusH + 4
    Me.Controls("txtInput").Width = w - 2 * pad - sendW - 8
    Me.Controls("txtInput").Height = inputH
    
    Me.Controls("btnSend").Left = Me.Controls("txtInput").Left + Me.Controls("txtInput").Width + 8
    Me.Controls("btnSend").Top = Me.Controls("txtInput").Top
    Me.Controls("btnSend").Width = sendW
    Me.Controls("btnSend").Height = inputH
    
    Me.Controls("btnStop").Left = Me.Controls("btnSend").Left
    Me.Controls("btnStop").Top = Me.Controls("btnSend").Top
    Me.Controls("btnStop").Width = sendW
    Me.Controls("btnStop").Height = inputH
    
    Me.Controls("btnNew").Left = pad
    Me.Controls("btnNew").Top = Me.Controls("txtInput").Top + inputH + 4
    Me.Controls("btnClear").Left = Me.Controls("btnNew").Left + Me.Controls("btnNew").Width + 8
    Me.Controls("btnClear").Top = Me.Controls("btnNew").Top
    Me.Controls("btnContinue").Left = Me.Controls("btnClear").Left + Me.Controls("btnClear").Width + 8
    Me.Controls("btnContinue").Top = Me.Controls("btnNew").Top
    On Error GoTo 0
End Sub

Public Sub btnSend_Click()
    SendMessage
End Sub

Public Sub btnStop_Click()
    CancelConversation
End Sub

Public Sub btnContinue_Click()
    ContinueFromStop
End Sub

Public Sub btnNew_Click()
    ResetConversation
    Me.Controls("txtChat").Text = ""
    m_conversationStarted = False
    Me.Controls("btnContinue").Visible = False
    AppendChat "Z.AI", T("chat.new_started")
    SetStatus T("chat.ready")
End Sub

Public Sub btnClear_Click()
    Me.Controls("txtChat").Text = ""
End Sub

Public Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        KeyCode = 0
        SendMessage
    End If
End Sub

Private Sub SendMessage()
    Dim txtInput As MSForms.TextBox
    Set txtInput = Me.Controls("txtInput")
    
    Dim userText As String
    userText = Trim(txtInput.Text)
    If userText = "" Then Exit Sub
    
    If Not m_conversationStarted Then
        InitConversation
        m_conversationStarted = True
    End If
    
    AppendChat "Ty", userText
    txtInput.Text = ""
    
    SetProcessing True
    DoEvents
    
    Dim response As String
    response = SendUserMessage(userText)
    
    AppendChat "Z.AI", response
    SetProcessing False
    txtInput.SetFocus
End Sub

Private Sub ContinueFromStop()
    If IsProcessing() Then Exit Sub
    If Not CanContinue() Then Exit Sub
    
    Me.Controls("btnContinue").Visible = False
    AppendChat "Z.AI", T("chat.continuing")
    
    SetProcessing True
    DoEvents
    
    Dim response As String
    response = ContinueConversation()
    
    AppendChat "Z.AI", response
    SetProcessing False
    Me.Controls("txtInput").SetFocus
End Sub

Private Sub SetProcessing(ByVal processing As Boolean)
    On Error Resume Next
    Me.Controls("txtInput").Enabled = Not processing
    Me.Controls("btnSend").Visible = Not processing
    Me.Controls("btnSend").Enabled = Not processing
    Me.Controls("btnStop").Visible = processing
    If processing Then
        SetStatus T("chat.processing")
        Me.Controls("btnContinue").Visible = False
    Else
        SetStatus TFormat("chat.ready_count", GetMessageCount())
        Me.Controls("btnContinue").Visible = CanContinue()
    End If
    DoEvents
    On Error GoTo 0
End Sub

Private Sub AppendChat(ByVal sender As String, ByVal message As String)
    On Error Resume Next
    
    Dim txtChat As MSForms.TextBox
    Set txtChat = Me.Controls("txtChat")
    
    Dim timestamp As String
    timestamp = Format(Now, "hh:nn")
    
    Dim newText As String
    If txtChat.Text <> "" Then newText = txtChat.Text & vbCrLf & vbCrLf
    newText = newText & "[" & timestamp & "] " & sender & ":" & vbCrLf & message
    
    txtChat.Text = newText
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelLength = 0
    
    On Error GoTo 0
End Sub

Private Sub SetStatus(ByVal text As String)
    On Error Resume Next
    Me.Controls("lblStatus").Caption = text
    DoEvents
    On Error GoTo 0
End Sub

Private Sub UserForm_Terminate()
    LogInfo "Chat form closed"
End Sub
