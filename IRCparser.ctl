VERSION 5.00
Begin VB.UserControl IRCparser 
   ClientHeight    =   1308
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1848
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1308
   ScaleWidth      =   1848
End
Attribute VB_Name = "IRCparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_usernick = "0"
Const m_def_userhost = "0"
Const m_def_msgtype = "0"
Const m_def_msgtext = "0"
Const m_def_chanName = "0"
'Property Variables:
Dim m_usernick As String
Dim m_userhost As String
Dim m_msgtype As String
Dim m_msgtext As String
Dim m_chanName As String


Public Property Get usernick() As String
Attribute usernick.VB_MemberFlags = "400"
    usernick = m_usernick
End Property

Public Property Let usernick(ByVal New_usernick As String)
    If Ambient.UserMode = False Then err.Raise 382
    If Ambient.UserMode Then err.Raise 393
    m_usernick = New_usernick
    PropertyChanged "usernick"
End Property

Public Property Get userhost() As String
Attribute userhost.VB_MemberFlags = "400"
    userhost = m_userhost
End Property

Public Property Let userhost(ByVal New_userhost As String)
    If Ambient.UserMode = False Then err.Raise 382
    If Ambient.UserMode Then err.Raise 393
    m_userhost = New_userhost
    PropertyChanged "userhost"
End Property

Public Property Get msgtype() As String
Attribute msgtype.VB_MemberFlags = "400"
    msgtype = m_msgtype
End Property

Public Property Let msgtype(ByVal New_msgtype As String)
    If Ambient.UserMode = False Then err.Raise 382
    If Ambient.UserMode Then err.Raise 393
    m_msgtype = New_msgtype
    PropertyChanged "msgtype"
End Property

Public Property Get msgtext() As String
Attribute msgtext.VB_MemberFlags = "400"
    msgtext = m_msgtext
End Property

Public Property Let msgtext(ByVal New_msgtext As String)
    If Ambient.UserMode = False Then err.Raise 382
    If Ambient.UserMode Then err.Raise 393
    m_msgtext = New_msgtext
    PropertyChanged "msgtext"
End Property

Public Function parse(text As String) As Long

End Function

Public Function tokenise(text As String) As Variant

End Function

Public Property Get chanName() As String
Attribute chanName.VB_MemberFlags = "400"
    chanName = m_chanName
End Property

Public Property Let chanName(ByVal New_chanName As String)
    If Ambient.UserMode = False Then err.Raise 382
    If Ambient.UserMode Then err.Raise 393
    m_chanName = New_chanName
    PropertyChanged "chanName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_usernick = m_def_usernick
    m_userhost = m_def_userhost
    m_msgtype = m_def_msgtype
    m_msgtext = m_def_msgtext
    m_chanName = m_def_chanName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_usernick = PropBag.ReadProperty("usernick", m_def_usernick)
    m_userhost = PropBag.ReadProperty("userhost", m_def_userhost)
    m_msgtype = PropBag.ReadProperty("msgtype", m_def_msgtype)
    m_msgtext = PropBag.ReadProperty("msgtext", m_def_msgtext)
    m_chanName = PropBag.ReadProperty("chanName", m_def_chanName)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("usernick", m_usernick, m_def_usernick)
    Call PropBag.WriteProperty("userhost", m_userhost, m_def_userhost)
    Call PropBag.WriteProperty("msgtype", m_msgtype, m_def_msgtype)
    Call PropBag.WriteProperty("msgtext", m_msgtext, m_def_msgtext)
    Call PropBag.WriteProperty("chanName", m_chanName, m_def_chanName)
End Sub

