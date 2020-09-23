VERSION 5.00
Object = "{FFACF7F3-B868-11CE-84A8-08005A9B23BD}#1.7#0"; "DSSOCK32.OCX"
Begin VB.Form Form1 
   Caption         =   "jimm_IRC version 1.0B"
   ClientHeight    =   4815
   ClientLeft      =   1125
   ClientTop       =   1605
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   Begin VB.ListBox NameList 
      Height          =   3960
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1485
   End
   Begin VB.TextBox Outgoing 
      Height          =   300
      Left            =   64
      TabIndex        =   1
      Top             =   3960
      Width           =   5220
   End
   Begin VB.TextBox Incoming 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Akuru-Bodu"
         Size            =   20.25
         Charset         =   2
         Weight          =   1084
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3588
      Left            =   64
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5232
   End
   Begin dsSocketLib.dsSocket ds1 
      Height          =   420
      Left            =   6360
      TabIndex        =   3
      Top             =   4335
      Width           =   420
      _Version        =   65543
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
      LocalPort       =   0
      RemoteHost      =   ""
      RemotePort      =   0
      ServiceName     =   ""
      RemoteDotAddr   =   ""
      Linger          =   -1  'True
      Timeout         =   10
      LineMode        =   0   'False
      EOLChar         =   10
      BindConnect     =   0   'False
      SocketType      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This is a test profram! so please don't complain! type /join #kotari...once u get online"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   5940
   End
   Begin VB.Label Topic 
      Caption         =   "Topic:"
      Height          =   192
      Left            =   84
      TabIndex        =   4
      Top             =   84
      Width           =   6684
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu FileConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu FileSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example and included document are
' Copyright (C) 1998 by sofwath

' Please read the document that comes with this
' program.

Dim CRLF As String ' Cairrage return/Line feed
Dim OldText As String ' Holds any text still
                      ' needing processed
Dim channel As String ' Holds the channel name
Dim CMode ' CurrentMode of client
          ' 0 is logged in
          ' 1 is joining channel
          ' 2 is in channel
          
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
      textmsg = text$
  Incoming.text = Incoming.text & textmsg & CRLF

End Sub


Sub SendData(textmsg As String)

  ' Send the data in textmsg to the server, and
  ' add a CRLF
  ds1.Send = textmsg & CRLF
    
End Sub

Private Sub ds1_Close(ErrorCode As Integer, ErrorDesc As String)

  FileConnect.Caption = "&Connect"
  AddText "*** Disconnected (" & Str$(ErrorCode) & ": " & ErrorDesc & ")"

End Sub

Private Sub ds1_Connect()

  ' Physical connect
  AddText "*** Connection established."
  AddText "*** Sending login information..."
  ' Send the server my nickname
  SendData "NICK " & Nickname
  ' Send the server the user information
  SendData "USER " & Nickname & " " & ds1.LocalName & " " & Server & " :Real Name"

End Sub

Private Sub ds1_Receive(ReceiveData As String)

  Dim inData As String
  Dim sline As String
  Dim msg As String
  Dim msg2 As String
  Dim x
  
  ' Get the incoming data into a string
  inData = ReceiveData
      
  ' Add any unprocessed text on first
  inData = OldText & inData
  
  ' Some IRC servers are only using a Cairrage
  ' Return, or a LineFeed, instead of both, so
  ' we need to be prepared for that
  x = 0
  If Right$(inData, 2) = CRLF Then x = 1
  If Right$(inData, 1) = Chr$(10) Then x = 1
  If Right$(inData, 1) = Chr$(13) Then x = 1
  If x = 1 Then
    OldText = "" ' its a full send, process
  Else
    OldText = inData: Exit Sub ' incomplete send
                               ' save and exit
  End If
  
again:
  GoSub parsemsg ' get next msg fragment
  'AddText sline
  
  If Left$(sline, 6) = "PING :" Then ' we need to pong to stay alive
    plen = Len(sline)
    getnum$ = Right$(sline, plen - 6)
    'MsgBox getnum$
    AddText "PING? PONG!"
    SendData "PONG " & getnum$ & Server
    GoTo again ' get next msg
  End If
  If Left$(sline, 5) = "ERROR" Then ' some error
    AddText "*** ERROR " & Mid$(sline, InStr(sline, "("))
  End If
  If Left$(sline, Len(Nickname) + 1) = ":" & Nickname Then
    ' a command for the client only
    sline = Mid$(sline, InStr(sline, " ") + 1)
    Select Case Left$(sline, InStr(sline, " ") - 1)
      Case "MODE"
        AddText "*** Your mode is now " & Mid$(sline, InStr(sline, ":") + 1)
    End Select
  End If
  If Mid$(sline, InStr(sline, " ") + 1, 7) = "PRIVMSG" Then
    'someone /msged us
    msg = Mid$(sline, InStr(sline, " ") + 9)
    If LCase$(Left$(msg, InStr(msg, " ") - 1)) = LCase$(Nickname) Then ' private msg
      ' add so its: --nick-- msg here
      AddText "--" & Mid$(sline, 2, InStr(sline, "!") - 2) & "-- " & Mid$(msg, InStr(msg, ":") + 1)
    End If
  End If
  Select Case CMode
    Case 0 ' not in channel
      If Mid$(sline, InStr(1, sline, " ") + 1, 3) = "001" Then
        Server = Mid$(sline, 2, InStr(sline, " ") - 2)
      End If
      If Mid$(sline, InStr(1, sline, " ") + 1, 3) = "433" Then
        AddText "* " & Nickname & " is already in use."
        FileConnect.Caption = "&Connect"
        AddText "*** Disconnected"
        ds1.Close: Exit Sub
      End If
      If Left$(sline, Len(Server) + 1) = ":" & Server Then
        ' its a server msg, add the important part
        sline = Mid$(sline, InStr(2, sline, ":") + 1)
        AddText sline
      End If
      If Left$(sline, 13) = "NOTICE AUTH :" Then
        sline = Mid$(sline, InStr(2, sline, ":") + 1)
        AddText sline
      End If
    Case 1 ' joining channel
      If Left$(sline, Len(Server) + 1) = ":" & Server Then
        msg = Mid$(sline, InStr(sline, " ") + 1)
        Select Case Left$(msg, InStr(msg, " ") - 1)
          Case "332" ' Topic
            Topic.Caption = "Topic: " & Mid$(msg, InStr(msg, ":") + 1)
          Case "353" ' Name list
            msg = Mid$(msg, InStr(msg, ":") + 1)
            Do Until msg = "" ' break apart names and add them seperatly
              x = InStr(msg, " ")
              If x <> 0 Then
                NameList.AddItem Left$(msg, x - 1)
                msg = Mid$(msg, x + 1)
              Else
                NameList.AddItem msg
                msg = ""
              End If
            Loop
          Case "366" ' End of Name List
            CMode = 2 ' change mode to joined channel
        End Select
      Else
        ' someone joined the channel, us!
        If Left$(sline, InStr(sline, " ") - 1) = "JOIN" Then
          AddText "*** " & Nickname & " has joined " & channel
        End If
      End If
    Case 2 ' in a channel
      If Mid$(sline, InStr(sline, " ") + 1, 7) = "PRIVMSG" Then
        msg = Mid$(sline, InStr(sline, " ") + 9)
        If LCase$(Left$(msg, InStr(msg, " ") - 1)) = LCase$(Nickname) Then ' private msg
          AddText "--" & Mid$(sline, 2, InStr(sline, "!") - 2) & "-- " & Mid$(msg, InStr(msg, ":") + 1)
        Else ' channel msg
          If Left$(Mid$(msg, InStr(msg, ":") + 1), 1) = Chr$(1) Then ' action
            msg2 = Mid$(msg, InStr(msg, ":") + 9)
            AddText "* " & Mid$(sline, 2, InStr(sline, "!") - 2) & " " & Left$(msg2, Len(msg2) - 1)
          Else ' msg
            AddText "<" & Mid$(sline, 2, InStr(sline, "!") - 2) & "> " & Mid$(msg, InStr(msg, ":") + 1)
          End If
        End If
      Else
        If Mid$(sline, InStr(sline, " ") + 1, 4) = "PART" Then
          msg = Mid$(sline, 2, InStr(sline, "!") - 2) ' nickname of user
          AddText "*** " & msg & " (" & Mid$(sline, InStr(sline, "!") + 1, InStr(InStr(sline, "!") + 1, sline, " ") - InStr(sline, "!") - 1) & ") has left " & Mid$(sline, InStr(sline, "PART") + 5)
          ' Cycle through the names list for their nickname, and remove it
          For x = 0 To NameList.ListCount
            ' Check if this index is their nick, case-insensitive
            If LCase$(NameList.List(x)) = LCase$(msg) Then
              NameList.RemoveItem x
              Exit For
            End If
          Next
          GoTo again
        End If
        If Mid$(sline, InStr(sline, " ") + 1, 4) = "JOIN" Then
          msg = Mid$(sline, 2, InStr(sline, "!") - 2) ' nickname of user
          AddText "*** " & msg & " (" & Mid$(sline, InStr(sline, "!") + 1, InStr(InStr(sline, "!") + 1, sline, " ") - InStr(sline, "!") - 1) & ") has joined " & Mid$(sline, InStr(2, sline, ":") + 1)
          NameList.AddItem msg
          GoTo again
        End If
        If Mid$(sline, InStr(sline, " ") + 1, 4) = "NICK" Then
          msg = Mid$(sline, 2, InStr(sline, "!") - 2) ' nickname of user
          AddText "*** " & msg & " is now known as " & Mid$(sline, InStr(2, sline, ":") + 1)
          ' Cycle through the names list for their nickname, and remove it
          For x = 0 To NameList.ListCount
            ' Check if this index is their nick, case-insensitive
            If LCase$(NameList.List(x)) = LCase$(msg) Then
              NameList.RemoveItem x
              Exit For
            End If
          Next
          ' Add the new nickname to the list
          NameList.AddItem Mid$(sline, InStr(2, sline, ":") + 1)
          GoTo again
        End If
        ' command not yet supported, just display it
        AddText sline
      End If
  End Select
  GoTo again
Exit Sub

parsemsg:
  ' IRC may send more than one msg at a time,
  ' so parse them first
  If inData = "" Then Exit Sub
  x = InStr(inData, CRLF) ' find the break
  If x <> 0 Then
    sline = Left$(inData, x - 1)
    ' strip off the text
    If Len(inData) > x + 2 Then
      inData = Mid$(inData, x + 2)
    Else
      inData = ""
    End If
  Else
    x = InStr(inData, Chr$(13)) ' find the break
    If x = 0 Then
      x = InStr(inData, Chr$(10)) ' find the break
    End If
    If x <> 0 Then
      sline = Left$(inData, x - 1)
    Else
      sline = inData
    End If
    ' strip off the text
    If Len(inData) > x + 1 Then
      inData = Mid$(inData, x + 1)
    Else
      inData = ""
    End If
  End If
Return

End Sub
Private Sub FileConnect_Click()

  If FileConnect.Caption = "&Connect" Then
    ' Clear textbox, topic and listbox
    Incoming.text = ""
    NameList.Clear
    Topic.Caption = "Topic: "
    AddText "*** Attempting to connect to " & Server & "..."
    Incoming.Refresh
    FileConnect.Caption = "&Disconnect"
    ' Set the RemoteHost to the IRC Server Host
    Server = setup.ServerCombo.text
    ds1.RemoteHost = Server
    ' Set the Port to connect to
    ds1.RemotePort = Port
    ' Connect
    ds1.Connect
  Else
    FileConnect.Caption = "&Connect"
    AddText "*** Disconnected"
    ' Close the socket
    ds1.Close
  End If

End Sub
Private Sub FileExit_Click()

  ' Close the IRC server connection
  If FileConnect.Caption = "&Disconnect" Then ds1.Close
  ' Close the program
  Unload Me

End Sub
Private Sub FileSetup_Click()

  ' Show the setup form
  setup.Show 1

End Sub
Private Sub Form_Activate()

  ' Scroll the textbox down again
  Incoming_Change
  
End Sub

Private Sub Form_Load()

  ' Set CRLF to be Cairrage Return + Line Feed,
  ' ALL IRC messages end with this
  CRLF = Chr$(13) & Chr$(10)
  ' Set the current mode to 0
  CMode = 0
  
  'Set the default values
  Server = "sandiego.ca.us.undernet.org"
  Port = 6667
  Nickname = "jim_irc"
  
End Sub

Private Sub HelpAbout_Click()

  about.Show 1

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




Private Sub Outgoing_KeyPress(KeyAscii As Integer)

  Dim msg As String
  
  ' Exit unless its a return, then process
  If KeyAscii <> 13 Then Exit Sub
  KeyAscii = 0 ' Stop that stupid beep!
  msg = Outgoing.text
  If Left$(msg, 1) <> "/" Then
    ' they want to send a msg, send it if we're
    ' in a channel
    If NameList.ListCount > 0 Then
      SendData "PRIVMSG " & channel & " :" & msg
      AddText "> " & msg
    End If
  Else
    Outgoing.text = Mid$(Outgoing.text, 2)
    msg = Mid$(Outgoing.text, InStr(Outgoing.text, " ") + 1)
    Select Case UCase$(Left$(Outgoing.text, InStr(Outgoing.text, " ") - 1)) ' see what kind of action to do
      Case "JOIN"
        SendData "JOIN " & msg: CMode = 1 ' join the channel, set the mode
        channel = msg
      Case "ME"
        ' if we're in a channel, then do an action
        If NameList.ListCount > 0 Then SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION " & msg & Chr$(1)
        AddText "* " & Nickname & " " & msg
      Case "MSG"
        ' send a priv msg
        SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :" & Mid$(msg, InStr(msg, " ") + 1)
        AddText "=->" & Left$(msg, InStr(msg, " ") - 1) & "<-= " & Mid$(msg, InStr(msg, " ") + 1)
      Case Else
        MsgBox "sorry! this poor clinet does not know all those commands! <--teach me!"
    End Select
  End If
  ' clear the textbox
  Outgoing.text = ""

End Sub





