VERSION 5.00
Begin VB.Form setup 
   Caption         =   "KotarIRC Setup"
   ClientHeight    =   4935
   ClientLeft      =   1575
   ClientTop       =   1425
   ClientWidth     =   7335
   Icon            =   "setup.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4935
   ScaleWidth      =   7335
   Begin VB.Frame serverSetup 
      Caption         =   "Server Setup"
      Height          =   2460
      Left            =   4188
      TabIndex        =   18
      Top             =   108
      Width           =   3084
      Begin VB.TextBox serupIRC 
         Height          =   288
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   2028
         Width           =   1400
      End
      Begin VB.TextBox setupIRC 
         Height          =   288
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   1284
         Width           =   2856
      End
      Begin VB.TextBox setupIRC 
         Height          =   288
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   612
         Width           =   2856
      End
      Begin VB.Label paswdLabel 
         Caption         =   "Server Password: (for special users)"
         Height          =   252
         Index           =   2
         Left            =   96
         TabIndex        =   24
         Top             =   1680
         Width           =   2844
      End
      Begin VB.Label IPlabel 
         Caption         =   "Server Ports:"
         Height          =   252
         Index           =   1
         Left            =   96
         TabIndex        =   21
         Top             =   1032
         Width           =   1000
      End
      Begin VB.Label IPlabel 
         Caption         =   "Server IP:"
         Height          =   252
         Index           =   0
         Left            =   96
         TabIndex        =   19
         Top             =   360
         Width           =   768
      End
   End
   Begin VB.Frame setupIdentd 
      Caption         =   "Identd Setup"
      Height          =   1776
      Left            =   96
      TabIndex        =   11
      Top             =   2640
      Width           =   3960
      Begin VB.TextBox identPort 
         Height          =   288
         Left            =   900
         TabIndex        =   17
         Top             =   1212
         Width           =   624
      End
      Begin VB.TextBox sysType 
         Height          =   288
         Left            =   900
         TabIndex        =   16
         Top             =   756
         Width           =   840
      End
      Begin VB.TextBox identUser 
         Height          =   288
         Left            =   900
         TabIndex        =   12
         Top             =   312
         Width           =   1500
      End
      Begin VB.Label setupLabels 
         Caption         =   "Port:"
         Height          =   252
         Index           =   6
         Left            =   96
         TabIndex        =   15
         Top             =   1284
         Width           =   612
      End
      Begin VB.Label setupLabels 
         Caption         =   "OS Type:"
         Height          =   252
         Index           =   5
         Left            =   96
         TabIndex        =   14
         Top             =   816
         Width           =   696
      End
      Begin VB.Label setupLabels 
         Caption         =   "User ID:"
         Height          =   252
         Index           =   4
         Left            =   96
         TabIndex        =   13
         Top             =   400
         Width           =   612
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   5088
      TabIndex        =   1
      Top             =   4488
      Width           =   1040
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6228
      TabIndex        =   0
      Top             =   4488
      Width           =   1040
   End
   Begin VB.Frame IRCFrame 
      Caption         =   "IRC Setup"
      Height          =   2460
      Left            =   96
      TabIndex        =   2
      Top             =   108
      Width           =   3960
      Begin VB.TextBox fullnametext 
         Height          =   288
         Left            =   900
         TabIndex        =   9
         Top             =   1812
         Width           =   2856
      End
      Begin VB.TextBox NickText 
         Height          =   288
         Left            =   900
         TabIndex        =   5
         Top             =   1272
         Width           =   1920
      End
      Begin VB.TextBox PortText 
         Height          =   312
         Left            =   900
         TabIndex        =   4
         Text            =   "6667"
         Top             =   768
         Width           =   636
      End
      Begin VB.ComboBox ServerCombo 
         Height          =   288
         ItemData        =   "setup.frx":030A
         Left            =   900
         List            =   "setup.frx":030C
         TabIndex        =   3
         Top             =   324
         Width           =   2856
      End
      Begin VB.Label setupLabels 
         Caption         =   "Full Name:"
         Height          =   252
         Index           =   3
         Left            =   96
         TabIndex        =   10
         Top             =   1908
         Width           =   768
      End
      Begin VB.Label setupLabels 
         Caption         =   "Nickname:"
         Height          =   204
         Index           =   2
         Left            =   96
         TabIndex        =   8
         Top             =   1368
         Width           =   750
      End
      Begin VB.Label setupLabels 
         Caption         =   "Port:"
         Height          =   252
         Index           =   1
         Left            =   96
         TabIndex        =   7
         Top             =   828
         Width           =   612
      End
      Begin VB.Label setupLabels 
         Caption         =   "Server:"
         Height          =   252
         Index           =   0
         Left            =   96
         TabIndex        =   6
         Top             =   384
         Width           =   612
      End
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()

  ' Close setup
  Me.Hide
  
End Sub


Private Sub Form_Load()
    Dim i, upper As Integer
    

  ' Center myself in the middle of the screen
  Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
  Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
  ' get the current setting from the ini file
  identUser.text = userid
  sysType.text = ostype
  identPort.text = idport
  NickText.text = Nickname
  fullnametext.text = username
  setupIRC(0).text = server
  setupIRC(1).text = port
  ServerCombo.SelText = server 'ServerCombo.List(i - 1)
  
End Sub

Private Sub OK_Click()
    Dim s&
    Dim section, filename As String
    section = "irc setup"
    filename = App.Path & "\user.ini"
    
  ' Make sure all fields have data
  'If ServerCombo.text = "" Then
  '  Beep
  '  ServerCombo.SetFocus: Exit Sub
  'End If
  'If PortText.text = "" Then
  '  Beep
  '  PortText.SetFocus: Exit Sub
  'End If
  'If NickText.text = "" Then
  '  Beep
  '  NickText.SetFocus: Exit Sub
  'End If

  ' Set the global variables
'----------------------------------
    'writing changes to user.ini file
  s& = writeUserINI(ByVal section, ByVal "server", setupIRC(0).text)
  s& = writeUserINI(ByVal section, ByVal "realname", fullnametext.text)
  s& = writeUserINI(ByVal section, ByVal "nickname", NickText.text)
  s& = writeUserINI(ByVal section, ByVal "userid", identUser.text)
  s& = writeUserINI(ByVal section, ByVal "ostype", sysType.text)
  s& = writeUserINI(ByVal section, ByVal "port", identPort.text)
  server = setupIRC(0).text
  port = PortText.text
  username = fullnametext.text
  Nickname = NickText.text
  ' Close setup
  Me.Hide
  

End Sub

'for selecting a server and displaying it's details
'in the setupIRC stuff - simon

Private Sub ServerCombo_Click()
    Dim pos As Integer
    pos = ServerCombo.ListIndex
    
    setupIRC(0).text = servers(pos + 1).IP
    setupIRC(1).text = servers(pos + 1).port
    
End Sub

