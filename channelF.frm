VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form channelF 
   Caption         =   "Channel "
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "channelF.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   7860
   Begin VB.ListBox Namelist 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   1524
   End
   Begin VB.TextBox Outgoing 
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   1440
      TabIndex        =   0
      Top             =   5640
      Width           =   6108
   End
   Begin RichTextLib.RichTextBox Incoming 
      Height          =   5175
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9128
      _Version        =   327680
      Enabled         =   -1  'True
      RightMargin     =   1
      TextRTF         =   $"channelF.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Akuru"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "channelF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a 4 use of #kotari
' Copyright (C) 1998 by sofwath, Simon
Private chanName As String ' stores the name for this channel (local)
Sub AddText(textmsg As String, Optional Kula)
    Dim temptext As String
  ' Add the data in textmsg to the Incoming
  ' text box and force the text down
  ' ahaha ! don't forget to convert the text to dhivehi!
      'MsgBox Language
      'If Language = lngThaana  And ConvertOnot = True Then
      If chanState(getChanIndex(Me.channelName)).Lang = lngthaana Then
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
        temptext = textmsg
        textmsg = text$ ' if the langugae is thana
        'Debug.Print
        Debug.Print text$
      End If
      
    Incoming.SelStart = Len(Incoming.text) 'move to insertion point

    If chanState(getChanIndex(Me.channelName)).Lang = lngthaana Then
        Incoming.SelAlignment = 1
        Incoming.SelFontName = "Akuru"
        Incoming.SelFontSize = 18
    Else 'english
        Incoming.SelAlignment = 0
        Incoming.SelFontName = "Arial"
        Incoming.SelFontSize = 10
    End If
    
    'set color to black of no color has been specified
    If IsMissing(Kula) Then Kula = QBColor(0)
    
    Incoming.SelColor = Kula 'set colot
    Incoming.SelText = textmsg & CRLF  'put in the message

  

End Sub



Private Sub Form_Activate()
If chanState(getChanIndex(Me.channelName)).Lang = lngEnglish Then
    ClientMain.mnuLangEnglish.Checked = True
    ClientMain.mnuLangThaana.Checked = False
Else
    ClientMain.mnuLangEnglish.Checked = False
    ClientMain.mnuLangThaana.Checked = True
End If

End Sub

'Private Sub Form_Activate()

  ' Scroll the textbox down again
 ' Incoming_Change
  
'End Sub

Private Sub Form_Load()
    Namelist.Move 0, 0, 1433, (ScaleHeight - Outgoing.Height)
    Incoming.Move 1433, 0, (ScaleWidth - 1433), (ScaleHeight - Outgoing.Height)
    Outgoing.Move 0, Incoming.Height, ScaleWidth
    'MsgBox chanName
    
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then   '--Echo
        Namelist.Move 0, 0, 1433, (ScaleHeight - Outgoing.Height)
        Incoming.Move 1433, 0, (ScaleWidth - 1433), Namelist.Height
        Outgoing.Move 0, Namelist.Height, ScaleWidth
    End If

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




Private Sub Namelist_DblClick()
'---Echo
On Error GoTo Errhandler

    Dim tempNick As String
    Dim ArrayCount As Integer
    Dim WindowExists As Boolean
    Dim index As Integer
    index = ClientMain.ActiveForm.Namelist.ListIndex
    tempNick = ClientMain.ActiveForm.Namelist.List(index)
    'cut the "@" n "+" , if infront of nick
    If Left$(tempNick, 1) = "@" Or Left$(tempNick, 1) = "+" Then
        tempNick = Right$(tempNick, Len(tempNick) - 1)
    End If
    If Not DoesPvtExists(tempNick) Then 'if no curren window exists
        CreatePvtWindow tempNick    'create a new one
    Else
        PvtWindows(GetPvtIndex(tempNick)).WindowState = vbNormal 'yeah need to open!
    End If
    
Exit Sub
Errhandler:
If VBA.Information.err = 9 Then
    Resume Next
End If


End Sub


Private Sub Namelist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'display popmenu if right button -Echo
If Button = vbRightButton Then
    PopupMenu ClientMain.nickControl
End If
End Sub

Private Sub Outgoing_KeyPress(KeyAscii As Integer)
  Dim msg As String
  Dim Actions As String
  ' Exit unless its a return, then process
  ' - Sofwath
  If KeyAscii = "11" Then 'control + k pressed
    msg = "" ' color
    Outgoing.text = ""
    Exit Sub
  End If
  If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0 ' Stop that stupid beep!
    msg = Outgoing.text
    'MsgBox chanState(Me.Tag).name
    channel = chanState(Me.Tag).name
    'channel = Trim(Mid$(Outgoing.text, InStr(Outgoing.text, " ") + 1))
    If Left$(msg, 1) <> "/" Then
      ' they want to send a msg, send it if we're
      ' in a channel
      ' MsgBox chanState(1).name
      ClientMain.SendData "PRIVMSG " & channel & " :" & msg
      ConvertOnot = True
      AddText "<" & Nickname & "> " & msg
    Else
      ConvertOnot = False
      Outgoing.text = Mid$(Outgoing.text, 2)
      msg = Mid$(Outgoing.text, InStr(Outgoing.text, " ") + 1)
      Actions = UCase$(Left$(Outgoing.text, InStr(Outgoing.text, " ") - 1)) ' see what kind of action to do
      ExecuteAction Actions, msg, Me '<<whats in here have been transered to this subrotine so we can share it. -echo
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
            Namelist.AddItem Left$(names, X - 1)
            names = Mid$(names, X + 1)
        Else
            Namelist.AddItem names
            names = ""
        End If
    Loop

End Sub

'Removes the user who has just departed the channel from the
'list. - simon
Sub RemoveNick(person As String)
    For X = 0 To Namelist.ListCount
        ' Check if this index is their nick, case-insensitive
        If LCase$(Namelist.List(X)) = LCase$(person) Then
            Namelist.RemoveItem X
            Exit For
        End If
    Next
End Sub

Sub setChanName(name As String)
    chanName = name
End Sub

Function channelName() As String
    channelName = chanName
End Function
