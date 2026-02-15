VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZaiChat
   Caption         =   "Z.AI - Asystent Excel"
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

' --- Form Initialize ---
Private Sub UserForm_Initialize()
    ' Setup will be called from build script or manually
    Me.Caption = T("chat.title")
    Me.Width = 520
    Me.Height = 620
    
    ' Create controls programmatically
    CreateControls
    
    m_conversationStarted = False
    
    ' Show welcome message
    AppendChat "Z.AI", T("chat.welcome")
    
    LogInfo "Chat form opened"
End Sub

' --- Create UI controls programmatically ---
Private Sub CreateControls()
    On Error Resume Next
    
    ' Chat history textbox
    Dim txtChat As MSForms.TextBox
    Set txtChat = Me.Controls.Add("Forms.TextBox.1", "txtChat")
    With txtChat
        .Left = 10
        .Top = 10
        .Width = 490
        .Height = 470
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .WordWrap = True
        .Font.Name = "Consolas"
        .Font.Size = 10
        .BackColor = &H00FFFFFF
    End With
    
    ' Status label
    Dim lblStatus As MSForms.Label
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Left = 10
        .Top = 490
        .Width = 490
        .Height = 18
        .Caption = T("chat.ready")
        .Font.Size = 8
        .ForeColor = &H00808080
    End With
    
    ' Input textbox
    Dim txtInput As MSForms.TextBox
    Set txtInput = Me.Controls.Add("Forms.TextBox.1", "txtInput")
    With txtInput
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
    Dim btnSend As MSForms.CommandButton
    Set btnSend = Me.Controls.Add("Forms.CommandButton.1", "btnSend")
    With btnSend
        .Left = 418
        .Top = 512
        .Width = 82
        .Height = 50
        .Caption = T("chat.send")
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' New conversation button
    Dim btnNew As MSForms.CommandButton
    Set btnNew = Me.Controls.Add("Forms.CommandButton.1", "btnNew")
    With btnNew
        .Left = 10
        .Top = 570
        .Width = 120
        .Height = 25
        .Caption = T("chat.new")
        .Font.Size = 8
    End With
    
    ' Clear button
    Dim btnClear As MSForms.CommandButton
    Set btnClear = Me.Controls.Add("Forms.CommandButton.1", "btnClear")
    With btnClear
        .Left = 140
        .Top = 570
        .Width = 100
        .Height = 25
        .Caption = T("chat.clear")
        .Font.Size = 8
    End With
    
    On Error GoTo 0
End Sub

' --- Send button click ---
Public Sub btnSend_Click()
    SendMessage
End Sub

' --- New conversation ---
Public Sub btnNew_Click()
    ResetConversation
    Me.Controls("txtChat").Text = ""
    m_conversationStarted = False
    AppendChat "Z.AI", T("chat.new_started")
    SetStatus T("chat.ready")
End Sub

' --- Clear chat ---
Public Sub btnClear_Click()
    Me.Controls("txtChat").Text = ""
End Sub

' --- Handle Enter key in input ---
Public Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        KeyCode = 0
        SendMessage
    End If
End Sub

' --- Core send message function ---
Private Sub SendMessage()
    Dim txtInput As MSForms.TextBox
    Set txtInput = Me.Controls("txtInput")
    
    Dim userText As String
    userText = Trim(txtInput.Text)
    
    If userText = "" Then Exit Sub
    
    ' Initialize conversation if needed
    If Not m_conversationStarted Then
        InitConversation
        m_conversationStarted = True
    End If
    
    ' Show user message
    AppendChat "Ty", userText
    txtInput.Text = ""
    
    ' Disable input during processing
    txtInput.Enabled = False
    Me.Controls("btnSend").Enabled = False
    SetStatus T("chat.processing")
    DoEvents
    
    ' Send to agent
    Dim response As String
    response = SendUserMessage(userText)
    
    ' Show response
    AppendChat "Z.AI", response
    
    ' Re-enable input
    txtInput.Enabled = True
    Me.Controls("btnSend").Enabled = True
    SetStatus TFormat("chat.ready_count", GetMessageCount())
    txtInput.SetFocus
End Sub

' --- Append text to chat display ---
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
    
    ' Scroll to bottom
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelLength = 0
    
    On Error GoTo 0
End Sub

' --- Set status text ---
Private Sub SetStatus(ByVal text As String)
    On Error Resume Next
    Me.Controls("lblStatus").Caption = text
    DoEvents
    On Error GoTo 0
End Sub

' --- Form cleanup ---
Private Sub UserForm_Terminate()
    LogInfo "Chat form closed"
End Sub
