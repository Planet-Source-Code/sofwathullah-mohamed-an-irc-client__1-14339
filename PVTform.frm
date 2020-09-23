VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form PvtForm 
   Caption         =   "Private Window"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "PVTform.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   7740
   Begin RichTextLib.RichTextBox PvtTxt 
      Height          =   4410
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7779
      _Version        =   327680
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   1
      TextRTF         =   $"PVTform.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Akuru"
         Size            =   27.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox PvtInp 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   7575
   End
End
Attribute VB_Name = "PvtForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PvtLanguage As Languages 'hold the language of the pvt window
Sub AddText(textmsg As String, Optional Kula As Long)
    Dim temptext As String
  
    If PvtWindowState(Val(Me.Tag)).Lang = lngthaana Then
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
    End If
    
    PvtTxt.SelStart = Len(PvtTxt.text) 'move to insertion point

    'set  the alignemnt n font
    If PvtWindowState(Val(Me.Tag)).Lang = lngthaana Then
        PvtTxt.SelAlignment = 1
        PvtTxt.SelFontName = "Akuru"
        PvtTxt.SelFontSize = 18
    Else 'english
        PvtTxt.SelAlignment = 0
        PvtTxt.SelFontName = "Arial"
        PvtTxt.SelFontSize = 10
    End If
    
    'set color to black of no color has been specified
    If IsMissing(Kula) Then Kula = QBColor(0)
    
    PvtTxt.SelColor = Kula 'set colot
    PvtTxt.SelText = textmsg & CRLF  'put in the message


    
    
    
    

End Sub

Private Sub Form_Activate()
If PvtWindowState(GetPvtIndex(Me.Caption)).Lang = lngEnglish Then
    ClientMain.mnuLangEnglish.Checked = True
    ClientMain.mnuLangThaana.Checked = False
Else
    ClientMain.mnuLangEnglish.Checked = False
    ClientMain.mnuLangThaana.Checked = True
End If

End Sub

'need to use rich text box as the main output text box for formatting purposes
'i.e normal text box cannot be used... -- Echo

Private Sub Form_Load()
    PvtTxt.Move 0, 0, ScaleWidth, (ScaleHeight - 288)
    PvtInp.Move 0, PvtTxt.Height, ScaleWidth
    If Language = lngthaana Then
        Me.PvtTxt.SelAlignment = 1 'right aligned
        Me.PvtTxt.Font.name = "Akuru"
        Me.PvtTxt.Font.Size = 18
    Else
        Me.PvtTxt.SelAlignment = 0  'left aligned
        Me.PvtTxt.Font.name = "Arial"
        Me.PvtTxt.Font.Size = 12
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'mark the window as deleted, so that it can be recycled
    On Error Resume Next
    PvtWindowState(Val(Me.Tag)).deleted = True
End Sub


Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then 'can change onli if it is not minimized
    PvtTxt.Move 0, 0, ScaleWidth, (ScaleHeight - 288)
    PvtInp.Move 0, PvtTxt.Height, ScaleWidth
End If
End Sub


Private Sub PvtInp_KeyPress(KeyAscii As Integer)
  Dim msg As String
  Dim Actions As String
  ' Exit unless its a return, then process
  If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    If Len(PvtInp.text) > 0 Then 'check whether anything to send
        msg = PvtInp.text
        If Left$(msg, 1) <> "/" Then
          ' send the msg to a user
            ClientMain.SendData "PRIVMSG " & Me.Caption & " :" & msg
            AddText "<" & Nickname & "> " & msg, QBColor(0) 'black color
        'can use an elseif here to ignore those actions that cannot be performed on a pvt window
        ElseIf UCase$(Left$(msg, 6)) = "/CLEAR" Then
            PvtTxt.text = ""
        Else 'anoter action
            PvtInp.text = Mid$(PvtInp.text, 2)
            msg = Mid$(PvtInp.text, InStr(PvtInp.text, " ") + 1)
            Actions = UCase$(Left$(PvtInp.text, InStr(PvtInp.text, " ") - 1)) ' see what kind of action to do
            ExecuteAction Actions, msg, Me '<<whats in here have
        End If
        PvtInp.text = "" 'clear the input line
    Else
        Beep    'beep if try to send an empty string
    End If
End Sub


Private Sub PvtTxt_Change()
    PvtTxt.SelStart = Len(PvtTxt.text)
End Sub


Private Sub PvtTxt_GotFocus()
  PvtInp.SetFocus
End Sub


