VERSION 5.00
Begin VB.Form channelF 
   Caption         =   "KotariIRC version 1.0b"
   ClientHeight    =   6084
   ClientLeft      =   1128
   ClientTop       =   1320
   ClientWidth     =   7584
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6084
   ScaleWidth      =   7584
   Begin VB.TextBox Outgoing 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1224
      TabIndex        =   1
      Top             =   5496
      Width           =   6168
   End
   Begin VB.TextBox Incoming 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5098
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   324
      Width           =   6180
   End
   Begin VB.ListBox NameList 
      Appearance      =   0  'Flat
      Height          =   5400
      Left            =   12
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   504
      Width           =   1433
   End
End
Attribute VB_Name = "channelF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a 4 use of #kotari
' Copyright (C) 1998 by sofwath, Simon
Public chanName As String ' stores the name for this channel





          
Sub AddText(textmsg As String)

  ' Add the data in textmsg to the Incoming
  ' text box and force the text down
  ' ahaha ! don't forget to convert the text to dhivehi!
  xlate (textmsg)
      le = Len(text$)
        For r = 1 To le - 1
            If r = 1 Then
                re$ = Mid$(text$, le, 1)
                a$ = a$ + re$
            End If
            re$ = Mid$(text$, le - r, 1)
            a$ = a$ + re$
        Next
      text$ = a$
      'MsgBox Language
      If Language = "T" Then
        textmsg = text$ ' if the langugae is thana
      End If
  
  Incoming.text = Incoming.text & textmsg & CRLF

End Sub



'Private Sub Form_Activate()

  ' Scroll the textbox down again
 ' Incoming_Change
  
'End Sub

Private Sub Form_Load()
    NameList.Move 0, 0, 1433, (ScaleHeight - Outgoing.Height)
    Incoming.Move 1433, 0, (ScaleWidth - 1433), (ScaleHeight - Outgoing.Height)
    Outgoing.Move 0, Incoming.Height, ScaleWidth
    'MsgBox chanName
    
    
End Sub

Private Sub Form_Resize()
    NameList.Move 0, 0, 1433, (ScaleHeight - Outgoing.Height)
    Incoming.Move 1433, 0, (ScaleWidth - 1433), NameList.Height
    Outgoing.Move 0, NameList.Height, ScaleWidth

End Sub

Private Sub Incoming_Change()

' We want this box to scroll down automatically.

  Incoming.SelStart = Len(Incoming.text)

' What this does is says, make the start of my
' selected text the end of the entire text,
' which effectively scrolls down the textbox,
' but does not select anything. The len()
' command returns the length of characters of
' the text, in a number.

End Sub


Private Sub Incoming_GotFocus()

' We don't want the client to be able to edit
' the Incoming textbox.

  Outgoing.SetFocus

' This make it so the user cannot click inside
' the Incoming text box, but can still scroll it.
' It does this by giving another object the
' focus.

End Sub




Private Sub NameList_Click()
    PopupMenu ClientMain.nickControl
    
    
End Sub

Private Sub Outgoing_KeyPress(KeyAscii As Integer)

  Dim msg As String
  
  ' Exit unless its a return, then process
  If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0 ' Stop that stupid beep!
    msg = Outgoing.text
    'MsgBox chanState(Me.Tag).name
    channel = chanState(Me.Tag).name
    'channel = Trim(Mid$(Outgoing.text, InStr(Outgoing.text, " ") + 1))
    If Left$(msg, 1) <> "/" Then
      ' they want to send a msg, send it if we're
      ' in a channel
      
        'MsgBox chanState(1).name
        ClientMain.SendData "PRIVMSG " & channel & " :" & msg
        AddText "<" & Nickname & "> " & msg
    
    Else
      Outgoing.text = Mid$(Outgoing.text, 2)
      msg = Mid$(Outgoing.text, InStr(Outgoing.text, " ") + 1)
      actions = UCase$(Left$(Outgoing.text, InStr(Outgoing.text, " ") - 1)) ' see what kind of action to do
      Select Case actions
        Case "JOIN"
            ClientMain.SendData "JOIN " & msg: CMode = 1  ' join the channel, set the mode
            channel = Trim(msg)
            newChannel
                
           

        Case "ME"
          ' if we're in a channel, then do an action
          If NameList.ListCount > 0 Then ClientMain.SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION " & msg & Chr$(1)
          AddText "* " & Nickname & " " & msg
        Case "MSG"
          ' send a priv msg
          ClientMain.SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :" & Mid$(msg, InStr(msg, " ") + 1)
          AddText "=->" & Left$(msg, InStr(msg, " ") - 1) & "<-= " & Mid$(msg, InStr(msg, " ") + 1)
        Case "VERSION"
            ClientMain.SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :VERSION"
            AddText "[VERSION] " & Left$(msg, InStr(msg, " ") - 1)
        Case "NICK"
          ' change the nick
          ClientMain.SendData "NICK " & msg
          Nickname = msg
        Case "WHOIS"
          ' lets see who this guy is
          ClientMain.SendData "WHOIS " & server & " " & msg
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
  Outgoing.text = ""

End Sub

'Adds nicks to the nicklist. This way we modularise more
'and add to reusability..not to mention make code more readable :)
' - simon

Sub CreateNickList(names As String)
    Dim X As Integer ' to hold position of item to add
    Do Until names = "" ' break apart names and add them seperatly
        X = InStr(names, " ")
        If X <> 0 Then
            NameList.AddItem Left$(names, X - 1)
            names = Mid$(names, X + 1)
        Else
            NameList.AddItem names
            names = ""
        End If
    Loop

End Sub

'Removes the user who has just departed the channel from the
'list. - simon
Sub RemoveNick(person As String)
    For X = 0 To NameList.ListCount
        ' Check if this index is their nick, case-insensitive
        If LCase$(NameList.List(X)) = LCase$(person) Then
            NameList.RemoveItem X
            Exit For
        End If
    Next
End Sub


