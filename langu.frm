VERSION 5.00
Begin VB.Form options 
   Caption         =   "Language"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3015
   Icon            =   "langu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame langFrame 
      Caption         =   "Language"
      Height          =   1116
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1392
      Begin VB.OptionButton Option1 
         Caption         =   "English"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   264
         Width           =   864
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Thaana"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   636
         Width           =   936
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1188
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'just changes the default language -- Echo
If Option1.value = True Then        'Change Language to english
    Language = lngEnglish
Else                                'Change Language to thana
    Language = lngthaana
End If
'Hide the form
    Unload Me
End Sub


Private Sub Command2_Click()
' Hide the form
  Unload Me
End Sub

Private Sub Form_Load()
'clicks in the current language
If Language = lngEnglish Then
    Option1.value = True
Else
    Option2.value = True
End If
End Sub


