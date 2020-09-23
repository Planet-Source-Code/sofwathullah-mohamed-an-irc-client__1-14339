VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form comForm 
   Caption         =   "Server Window"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "comForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   7740
   Begin RichTextLib.RichTextBox serverTxt 
      Height          =   4404
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   7236
      _ExtentX        =   12753
      _ExtentY        =   7779
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      RightMargin     =   1
      TextRTF         =   $"comForm.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox serverInp 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Top             =   4812
      Width           =   7692
   End
End
Attribute VB_Name = "comForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    serverTxt.Move 0, 0, ScaleWidth, (ScaleHeight - 288)
    serverInp.Move 0, serverTxt.Height, ScaleWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'dont let the user close the server window
'can be closed by the program itself!
'--Echo
If UnloadMode = vbFormControlMenu Then
    Cancel = True
End If
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then '--Echo
        serverTxt.Move 0, 0, ScaleWidth, (ScaleHeight - 288)
        serverInp.Move 0, serverTxt.Height, ScaleWidth
    End If
End Sub

Private Sub serverInp_KeyPress(KeyAscii As Integer)

  Dim msg As String
  
  ' Exit unless its a return, then process
  If KeyAscii <> 13 Then Exit Sub
  KeyAscii = 0 ' Stop that stupid beep!
  msg = serverInp.text
  
  If Left$(msg, 1) <> "/" Then
    'they want to send a msg, send it if we're
    ' in a channel
    'If NameList.ListCount > 0 Then
      ClientMain.SendData "PRIVMSG " & channel & " :" & msg
      channels(getChanIndex(channel)).AddText "> " & msg
    'End If
  Else
    serverInp.text = Mid$(serverInp.text, 2)
    
    msg = Mid$(serverInp.text, InStr(serverInp.text, " ") + 1)
    
    Actions = UCase$(Left$(serverInp.text, InStr(serverInp.text, " ") - 1)) ' see what kind of action to do
    Select Case Actions
      Case "JOIN"
        channel = Trim(msg)
        ClientMain.SendData "JOIN " & msg: CMode = 1  ' join the channel, set the mode
        If chanCount = 1 Then
            Show
            ReDim channels(chanCount)
            ReDim chanState(chanCount)
            chanState(chanCount).name = channel
            chanState(chanCount).deleted = False
            chanState(chanCount).Lang = Language 'change to default language --Echo
            channels(chanCount).Tag = chanCount
            channels(chanCount).Caption = channel
            channels(chanCount).setChanName (channel)
            
            chanCount = chanCount + 1
            'MsgBox channels(1).name
        Else
            newChannel
            
        End If
        
        
      'Case "ME"
        ' if we're in a channel, then do an action
       ' If NameList.ListCount > 0 Then SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION " & msg & Chr$(1)
       ' AddText "* " & Nickname & " " & msg
      Case "MSG"
        ' send a priv msg
        ClientMain.SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :" & Mid$(msg, InStr(msg, " ") + 1)
        ClientMain.AddText "=->" & Left$(msg, InStr(msg, " ") - 1) & "<-= " & Mid$(msg, InStr(msg, " ") + 1), setcolor(Black)
      'Case "NICK"
        ' change the nick
      '  SendData "NICK " & msg
       ' Nickname = msg
      'Case "WHOIS"
        ' lets see who this guy is
       ' SendData "WHOIS " & server & " " & msg
      'Case "CTCP"
        'ClientMain.SendData
      Case "QUIT"
        ' end the IRC session with a message
        ClientMain.SendData "QUIT " & msg
      Case "PING"
        ' lets send a ping 2 the given address , could be a nick, or a server
        ClientMain.SendData "PING " & msg
      Case "HELP"
        ' show this dude the help
        Form3.Show 1
      Case Else
        MsgBox "sorry! this poor clinet does not know all those commands! <--teach me!"
    End Select
  End If
  ' clear the textbox
  serverInp.text = ""


End Sub

Private Sub serverTxt_Change()
    serverTxt.SelStart = Len(serverTxt.text)
End Sub

Private Sub serverTxt_GotFocus()
    serverInp.SetFocus
End Sub
