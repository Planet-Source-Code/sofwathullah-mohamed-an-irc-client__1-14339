VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm ClientMain 
   BackColor       =   &H8000000C&
   Caption         =   "KotarIRC '98"
   ClientHeight    =   5475
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9810
   Icon            =   "mainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ctcp 
      Left            =   1560
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock identd 
      Left            =   1080
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock ds1 
      Left            =   600
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu FileConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu Lang 
         Caption         =   "&Language"
      End
      Begin VB.Menu FileSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "&Language"
      Begin VB.Menu mnuLangEnglish 
         Caption         =   "English"
      End
      Begin VB.Menu mnuLangThaana 
         Caption         =   "&Thaana"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile &Horizontaly"
      End
      Begin VB.Menu mnuTileVertcal 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu helping 
         Caption         =   "&Help"
      End
   End
   Begin VB.Menu nickControl 
      Caption         =   "Control"
      Visible         =   0   'False
      Begin VB.Menu nickWhois 
         Caption         =   "Whois?"
      End
      Begin VB.Menu nickSlap 
         Caption         =   "Slap with.."
         Begin VB.Menu slapKanneli 
            Caption         =   "..kanneli!"
         End
         Begin VB.Menu slapTrout 
            Caption         =   "..trout!"
         End
         Begin VB.Menu slapKawaabu 
            Caption         =   "..kuni kawaabu!"
         End
      End
   End
End
Attribute VB_Name = "ClientMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub mnuArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Sub mnuLangEnglish_Click()
'this will be affeceted if the server window caption is changes!!!!
If Left$(Me.ActiveForm.Caption, 1) = "#" Then 'a channel window
    chanState(getChanIndex(Me.ActiveForm.channelName)).Lang = lngEnglish
ElseIf Me.ActiveForm.Caption = "Server Window" Then
    'do nuthin
Else    'a pvt window
    PvtWindowState(GetPvtIndex(Me.ActiveForm.Caption)).Lang = lngEnglish
End If

mnuLangEnglish.Checked = True
mnuLangThaana.Checked = False
End Sub

Sub mnuLangThaana_Click()
'note: this will be affected if server window caption is changes!! --Echo
If Left$(Me.ActiveForm.Caption, 1) = "#" Then 'a channel window
    chanState(getChanIndex(Me.ActiveForm.channelName)).Lang = lngthaana
ElseIf Me.ActiveForm.Caption = "Server Window" Then
    'do nuthin
Else    'a pvt window
    PvtWindowState(GetPvtIndex(Me.ActiveForm.Caption)).Lang = lngthaana
End If
mnuLangEnglish.Checked = False
mnuLangThaana.Checked = True

End Sub


Private Sub mnuTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertcal_Click()
    Me.Arrange vbTileVertical
End Sub

Sub AddText(textmsg As String, Kula As Long)

  ' Add the data in textmsg to the Incoming
  ' text box and force the text down
  ' ahaha ! don't forget to  the text to dhivehi!
      'MsgBox Language
      If Language = lngthaana And ConvertOnot = True Then
        MsgBox ConverOnot
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
         textmsg = text$ ' if the langugae is thana
        'MsgBox text$
        Debug.Print text$
      End If
  'setting the colour to whatever it is that u want
  comForm.serverTxt.SelColor = Kula
  'comForm.serverTxt.text = comForm.serverTxt.text & textmsg & CRLF
  
  'this way the updates are faster and we get the colouring working
  'anyway - simon
  comForm.serverTxt.SelText = textmsg & CRLF

End Sub

Private Sub ds1_Close()
  FileConnect.Caption = "&Connect"
  AddText "*** Disconnected (" & Str$(errorCode) & ": " & ErrorDesc & ")", setcolor(Blue)

End Sub

Private Sub ds1_DataArrival(ByVal bytesTotal As Long)
  Dim aparser As New parser

  Dim inData As String
  Dim sline As String
  Dim msg As String
  Dim msg2 As String
  Dim PingReq As String
  Dim CurrentTime As String
  Dim errorCode As String ' to store the error code if any -simon
  Dim X
  
  ' Get the incoming data into a string
  ds1.GetData inData, vbString
  
  'Debug.Print inData
  ' Add any unprocessed text on first
  inData = OldText & inData
  
  ' Some IRC servers are only using a Cairrage
  ' Return, or a LineFeed, instead of both, so
  ' we need to be prepared for that
  X = 0
  ' cut down on those IF statements - simon
  If Right$(inData, 2) = CRLF Or Right$(inData, 1) = Chr$(10) Or Right$(inData, 1) = Chr$(13) Then
    X = 1
  End If
  
  If X = 1 Then
    OldText = "" ' its a full send, process
  Else
    OldText = inData
    Exit Sub ' incomplete send  ' save and exit
  End If
  
again:
  GoSub parsemsg ' get next msg fragment
  'AddText sline
  
  If Left$(sline, 6) = "PING :" Then ' we need to pong to stay alive
    ConvertOnot = False
    plen = Len(sline)
    getnum$ = Right$(sline, plen - 6)  'lets get the ping number
    AddText "PING? PONG!", setcolor(Green)
    SendData "PONG " & getnum$ & server
    GoTo again ' get next msg
 
  ElseIf Left$(sline, 5) = "ERROR" Then ' some error
    ConvertOnot = False
    AddText "*** ERROR " & Mid$(sline, InStr(sline, "(")), setcolor(red)
  End If
  If Left$(sline, Len(Nickname) + 1) = ":" & Nickname Then
    ConvertOnot = False
    ' a command for the client only
    sline = Mid$(sline, InStr(sline, " ") + 1)
    Select Case Left$(sline, InStr(sline, " ") - 1)
      Case "MODE"
        AddText "*** Your mode is now " & Mid$(sline, InStr(sline, ":") + 1), setcolor(Green)
    End Select
  End If
  If this_command(sline, "PRIVMSG") Then
    ConvertOnot = Ture
    'Debug.Print sline
    'someone /msged us
    msg = Mid$(sline, InStr(sline, " ") + 9)
    Dim tempNick, tempmsg As String
    tempmsg = Mid$(msg, InStr(msg, ":") + 1)
    tempNick = LCase$(Left$(msg, InStr(msg, " ") - 1))
    ' Check for CTCP's  - sofwath
    If Trim(tempmsg) = "VERSION" Then   ' private msg
      tempNick = (Left$(sline, InStr(sline, "!") - 1)) 'get the person who did the request
      tempNick = Right$(tempNick, Len(tempNick) - 1)
      ConvertOnot = False
      ' add so its: --nick-- msg here
      AddText "[" & Mid$(sline, 2, InStr(sline, "!") - 2) & "] VERSION", setcolor(red)
      'NOTICE ctcp
      'lets tell this guys what we are using
      ' hey but we need to put the version in a varibale (laterz)
      SendData "NOTICE " + tempNick + " :VERSION kotariIRC BetA   8,1(http://members.xoom.com/kotari/)"
    ElseIf Left$(Trim(tempmsg), 5) = "PING" Then
        'ping reply to user
        tempNick = (Left$(sline, InStr(sline, "!") - 1)) 'get the person who did the request
        tempNick = Right$(tempNick, Len(tempNick) - 1)
        PingReq = Mid$(sline, InStr(sline, "") - 1)
        'lets reply to ping
        SendData "NOTICE " + tempNick + " " + PingReq
    ElseIf Trim(tempmsg) = "FINGER" Then
        ' finger reply
        tempNick = (Left$(sline, InStr(sline, "!") - 1)) 'get the person who did the request
        tempNick = Right$(tempNick, Len(tempNick) - 1)
        CurrentTime = Now
        SendData "NOTICE " + tempNick + " :" + "FINGER kotariIRC_User@" + userid + ""
    ElseIf Trim(tempmsg) = "TIME" Then
        ' timer reply
        tempNick = (Left$(sline, InStr(sline, "!") - 1)) 'get the person who did the request
        tempNick = Right$(tempNick, Len(tempNick) - 1)
        CurrentTime = Now
        SendData "NOTICE " + tempNick + " :" + "TIME " + CurrentTime + ""
    ElseIf tempNick = LCase$(Nickname) Then
      ConvertOnot = False
      AddText "[" & Mid$(sline, 2, InStr(sline, "!") - 2) & "]" & Mid$(msg, InStr(msg, ":") + 1), setcolor(Green)
    End If
  End If
    'we get the error code in advance
    'so that repetiveness is minimised
    'and it looks nicer :) - simon
    errorCode = Mid$(sline, InStr(1, sline, " ") + 1, 3)
    Select Case CMode

    
    Case 0 ' not in channel
      If errorCode = "001" Then
        server = Mid$(sline, 2, InStr(sline, " ") - 2)
      'making this an ElseIf cuts decision time by half - simon
      ElseIf errorCode = "433" Then
        ConvertOnot = False
        AddText "* " & Nickname & " is already in use.", setcolor(Black)
        FileConnect.Caption = "&Connect"
        ConvertOnot = False
        AddText "*** Disconnected", setcolor(Blue)
        ds1.Close: Exit Sub
      End If
      If Left$(sline, Len(server) + 1) = ":" & server Then
        ' its a server msg, add the important part
        sline = Mid$(sline, InStr(2, sline, ":") + 1)
        ConvertOnot = False
        AddText sline, setcolor(Black)
      End If
      If Left$(sline, 13) = "NOTICE AUTH :" Then
        ConvertOnot = False
        sline = Mid$(sline, InStr(2, sline, ":") + 1)
        AddText sline, setcolor(Blue)
      End If
    Case 1 ' joining channel
      If Left$(sline, Len(server) + 1) = ":" & server Then
        msg = Mid$(sline, InStr(sline, " ") + 1)
        Select Case CInt(Left$(msg, InStr(msg, " ") - 1))
          Case err.RPL_TOPIC  ' Topic
            channels(getChanIndex(channel)).Caption = channel & " " & Mid$(msg, InStr(msg, ":") + 1)
          Case err.RPL_NAMREPLY  ' Name list
            msg = Mid$(msg, InStr(msg, ":") + 1)
            'here we get the ppl in the channel
            
            channels(getChanIndex(channel)).CreateNickList (msg) 'lets do the list baby! = simon
          Case err.RPL_ENDOFNAMES  ' End of Name List
            CMode = 2 ' change mode to joined channel
        End Select
      Else
        ' someone joined the channel, us!
        ' there is a something missing we need to update the user list dho?
        ' but i'll fix it
        If Left$(sline, InStr(sline, " ") - 1) = "JOIN" Then
            ConvertOnot = False
            channels(getChanIndex(channel)).AddText "*** " & Nickname & " has joined " & channel
        End If
      End If
    Case 2 ' in a channel
        Debug.Print ">>>" & sline
        aparser.parse (sline)
      If aparser.msgType = "PRIVMSG" Then
        msg = Mid$(sline, InStr(sline, " ") + 9)
        messages msg, sline

      Else
        If aparser.msgType = "PART" Then
          msg = getNick(sline) ' nickname of user/ this function is in actionsMod - simon
          ConvertOnot = False
          channels(getChanIndex(channel)).AddText "*** " & msg & " (" & Mid$(sline, InStr(sline, "!") + 1, InStr(InStr(sline, "!") + 1, sline, " ") - InStr(sline, "!") - 1) & ") has left " & Mid$(sline, InStr(sline, "PART") + 5)
          ' Cycle through the names list for their nickname, and remove it
          channels(getChanIndex(channel)).RemoveNick (msg)
          GoTo again ' i hate this goto's but don't know what was on my mind then! <-- we will remove this soon
        ElseIf aparser.msgType = "JOIN" Then
          ConvertOnot = False
          msg = aparser.usernick
          'what the hell is happening here man? - simon
          '(getChanIndex(channel)) <-- removed that and
          'channels(getChanIndex(channel)). <-- replaced this -sofwath
          ConvertOnot = False
          channels(getChanIndex(channel)).AddText "*** " & msg & " (" & Mid$(sline, InStr(sline, "!") + 1, InStr(InStr(sline, "!") + 1, sline, " ") - InStr(sline, "!") - 1) & ") has joined " & Mid$(sline, InStr(2, sline, ":") + 1)
          'channels.(getChanIndex(channel)).AddText "*** " & msg & " (" & Mid$(sline, InStr(sline, "!") + 1, InStr(InStr(sline, "!") + 1, sline, " ") - InStr(sline, "!") - 1) & ") has joined " & Mid$(sline, InStr(2, sline, ":") + 1)
          ConvertOnot = False
          channels(getChanIndex(channel)).AddText msg
          ConvertOnot = False
          channels(getChanIndex(channel)).Namelist.AddItem msg
          GoTo again
        End If
        If this_command(sline, "NICK") Then 'nick change tha? -simon
          ConvertOnot = False
          msg = getNick(sline) ' nickname of user
          ConvertOnot = False
          channels(getChanIndex(channel)).AddText "*** " & msg & " is now known as " & Mid$(sline, InStr(2, sline, ":") + 1)
          ' Cycle through the names list for their nickname, and remove it
          channels(getChanIndex(channel)).RemoveNick (msg) 'remove that nick dude! -simon
          ' Add the new nickname to the list
          channels(getChanIndex(channel)).Namelist.AddItem Mid$(sline, InStr(2, sline, ":") + 1)
          'change the pvt window status for the user,,
          ChangePvtWindowNick msg, Mid$(sline, InStr(2, sline, ":") + 1)
          GoTo again
        End If
        ' command not yet supported, just display it
        channels(getChanIndex(channel)).AddText sline
      End If
  End Select
  GoTo again
Exit Sub

parsemsg:
      If inData = "" Then Exit Sub
      X = InStr(inData, CRLF) ' find the break
      If X <> 0 Then
        sline = Left$(inData, X - 1)
        ' strip off the text
        If Len(inData) > X + 2 Then
          inData = Mid$(inData, X + 2)
        Else
          inData = ""
        End If
      Else
        X = InStr(inData, Chr$(13)) ' find the break
        If X = 0 Then
          X = InStr(inData, Chr$(10)) ' find the break
        End If
        If X <> 0 Then
          sline = Left$(inData, X - 1)
        Else
          sline = inData
        End If
        ' strip off the text
        If Len(inData) > X + 1 Then
          inData = Mid$(inData, X + 1)
        Else
          inData = ""
        End If
      End If
Return

End Sub

Private Sub exit_Click()
    Unload Me
    End
End Sub



Private Sub identd_ConnectionRequest(ByVal requestID As Long)
    If identd.State <> sckClosed Then identd.Close
        identd.Accept requestID

End Sub

Private Sub identd_DataArrival(ByVal bytesTotal As Long)
    Dim idata As String
    identd.GetData idata, vbString
    
    AddText "[Identd request from " & identd.RemoteHostIP & " " & Left$(idata, Len(idata) - 2) & "]", setcolor(Magenta)
    AddText "[" & Left$(idata, Len(idata) - 2) & " : USERID : " & ostype & " : " & userid & "]", setcolor(Magenta)
    
    identd.SendData Left$(idata, Len(idata) - 2) & " : USERID : " & ostype & " : " & userid
    identd.Close
    identd.Listen
End Sub

Private Sub identd_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'if the socket is in use close identd and let it go
    If Number = sckAddressInUse Then
        identd.Close
    End If
End Sub

Private Sub Lang_Click()
' display the language form
 options.Show 1
End Sub


Private Sub FileSetup_Click()

  ' Show the setup form
  setup.Show 1

End Sub

Private Sub FileConnect_Click()


    Dim localport As Long

    ds1.Close 'close it just incase
    localport = 0
    
  If FileConnect.Caption = "&Connect" Then
    ' Clear textbox, topic and listbox
    comForm.serverTxt.text = ""
    'Namelist.Clear
    'Topic.Caption = "--Topic: "
    AddText "*** Attempting to connect to " & server & "...", setcolor(Black)
    comForm.serverTxt.Refresh
    FileConnect.Caption = "&Disconnect"
    ' Set the RemoteHost to the IRC Server Host
  
    port = 6667 ' irc port 2 connect
    'MsgBox server
    'ctcp.localport = 0
    'ctcp.Listen

    'MsgBox ds1.LocalHostName
   
    ' Set the Port to connect to
    ds1.RemotePort = port
    ds1.localport = localport
    identd.localport = 113 'identd listens on this port
    On Error GoTo closeidentd
        identd.Listen
closeidentd:
    ' Connect
    ds1.Connect server, port
   'MsgBox "connection...proceeding"
    'stat.Caption = " (you are online .... Now type /join #kotari).. /HELP for help!"
  Else
    FileConnect.Caption = "&Connect"
    AddText "*** Disconnected", setcolor(red)
    ' Close the socket
    ds1.Close
  End If

End Sub
Private Sub FileExit_Click()

  ' Close the IRC server connection
  If FileConnect.Caption = "&Disconnect" Then ds1.Close
  ' Close the program
  Unload Me
  End

End Sub

Private Sub HelpAbout_Click()

  frmAbout.Show 1

End Sub

Private Sub helping_Click()
' display help screen
    Form3.Show 1
End Sub

Private Sub ds1_Connect()

  ' Physical connect
  AddText "*** Connection established.", setcolor(Blue)
  AddText "*** Sending login information...", setcolor(Blue)
  ' Send the server my nickname
  SendData ("NICK " & Nickname)
  ' Send the server the user information
  SendData ("USER " & Nickname & " " & ds1.LocalHostName & " " & server & " :Real Name")

End Sub


Sub SendData(textmsg As String)

  ' Send the data in textmsg to the server, and
  ' add a CRLF
  ds1.SendData (textmsg & CRLF)
    
End Sub



Private Sub MDIForm_Load()
    ChDir App.Path
'    Show
    ' Always set the working directory to the directory containing the application.
 '   ReDim channels(1)
'    ReDim FState(1)
 '   channels(1).Tag = 1
    chanCount = 1
    comForm.Show  ' Set CRLF to be Cairrage Return + Line Feed,
   
  ' ALL IRC messages end with this
  CRLF = Chr$(13) & Chr$(10)
  ' Set the current mode to 0
  CMode = 0
  
  'Set the default values
  Dim logg As String
  Dim n&
  n& = GetUserName(ByVal logg, Len(logg))
  'MsgBox logg
  'Default Language
  Language = lngthaana
  readUserINI
  readServerList
  Nickname = setup.NickText.text 'assign the nick at start
  
  

End Sub

Function getChan(ins As String) As String
    Dim temp As String
    Dim start As Integer
    Dim ending As Integer
    
    start = InStr(1, ins, "#")
    ending = InStr(start, ins, " ")
    temp = Trim(Mid$(ins, start, ending - start))
    
    'Debug.Print temp
    getChan = temp
End Function

'i don't think this function should be herre..
'but for the moment..let it be eh? :) - simon
Sub messages(mesg As String, sline As String)
     Dim msg2 As String
     Dim SenderNick As String
    
    If LCase$(Left$(mesg, InStr(mesg, " ") - 1)) = LCase$(Nickname) Then ' private msg
        SenderNick = Left(sline, InStr(sline, "!") - 1) 'these 2 statements extract the senders nick
        SenderNick = Mid(SenderNick, 2, Len(SenderNick) - 1) '--Echo
        WriteToPvtWindow SenderNick, mesg 'write text on pvt window --Echo
    Else ' channel msg
        channel = getChan(mesg)
        If Left$(Mid$(mesg, InStr(mesg, ":") + 1), 1) = Chr$(1) Then ' action
            msg2 = Mid$(mesg, InStr(mesg, ":") + 9)
            ConvertOnot = True 'yes we need this to be converted
            channels(getChanIndex(channel)).AddText "* " & Mid$(sline, 2, InStr(sline, "!") - 2) & " " & Left$(msg2, Len(msg2) - 1)
        Else ' msg
            ConvertOnot = True
            channels(getChanIndex(channel)).AddText "<" & Mid$(sline, 2, InStr(sline, "!") - 2) & "> " & Mid$(mesg, InStr(mesg, ":") + 1)
        End If
    End If
    PrivateMsg = 0
End Sub

Private Sub nickWhois_Click()
    sendActions ("whois")
End Sub

Private Sub slapKanneli_Click()
    sendActions ("kanneli")
End Sub

Private Sub slapKawaabu_Click()
    sendActions ("kuni")

End Sub

Private Sub slapTrout_Click()
    sendActions ("trout")
End Sub

Private Sub sendActions(act As String)
    Dim tempNick As String
    Dim index As Integer
    index = ClientMain.ActiveForm.Namelist.ListIndex
    tempNick = ClientMain.ActiveForm.Namelist.List(index)
    
    
    If act = "kanneli" Then
        SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION slaps " & tempNick & " around a bit with a large kanneli!" & Chr$(1)
        ClientMain.ActiveForm.AddText "* " & Nickname & " slaps " & tempNick & "around a bit with a large kanneli!" & Chr$(1)
    ElseIf act = "trout" Then
        SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION slaps " & tempNick & " around a bit with a large trout!" & Chr$(1)
        ClientMain.ActiveForm.AddText "* " & Nickname & " slaps " & tempNick & "around a bit with a large trout!" & Chr$(1)
    ElseIf act = "kuni" Then
        SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION slaps " & tempNick & " around a bit with a kuni kawaabu!" & Chr$(1)
        ClientMain.ActiveForm.AddText "* " & Nickname & " slaps " & tempNick & "around a bit with a kuni kawaabu!" & Chr$(1)
    ElseIf act = "whois" Then
        If Left$(tempNick, 1) = "@" Or Left$(tempNick, 1) = "+" Then
            tempNick = Right$(tempNick, Len(tempNick) - 1)
        End If
        SendData "WHOIS " & server & " " & tempNick
    End If
End Sub


