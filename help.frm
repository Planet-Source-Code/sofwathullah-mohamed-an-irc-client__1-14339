VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Help"
   ClientHeight    =   4548
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6228
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4548
   ScaleWidth      =   6228
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4080
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' close the window
     Unload Me
End Sub

Private Sub Form_Load()
' help screen  .. just a smiple help
    Dim CRLF As String ' Cairrage return/Line feed
    Dim helping(13)
    CRLF = Chr$(13) & Chr$(10)
    Text1.text = " "
    helping(1) = " This is just a simple help screen!"
    helping(2) = "-----------------------------------"
    helping(3) = "Commands"
    helping(4) = "      type /join #<channel> to join a channel"
    helping(5) = "      type /msg <nick> <msg> to send a message to some one"
    helping(6) = "      type /who <text> to do a /me action (what ever it is)"
    helping(7) = "      type /whois <nick> to get the info on user"
    helping(8) = "      type /nick <nick> to change the nick"
    helping(9) = "      type /quit <msg> to end the current IRC session"
    helping(10) = "Language Selection "
    helping(11) = "      To change the Language click File and click on Language"
    '--------------
    For disp = 1 To 11
        Text1.text = Text1.text & helping(disp) & CRLF
    Next
End Sub
