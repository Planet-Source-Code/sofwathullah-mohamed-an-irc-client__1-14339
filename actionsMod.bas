Attribute VB_Name = "actionsMod"
Option Explicit
Public Enum Languages 'the options for language.. this way gonna be easier -Echo
    lngEnglish
    lngThaana
End Enum

' User-defined type to store information about child forms - simon
Type whichChannel       '<< used by pvt windows also --Echo
    name As String
    deleted As Boolean
    Lang As Languages   ' hold the language of the child form.. so can use diff languages for diff forms
End Type

'defining somethings common to servers
'so that we can store server configs - simon
Type serverType
    name As String
    IP As String
    port As String
End Type

'enumerating the colour names so that its easier to
'use in the function setcolour() below - simon
Enum colour
    Black
    Blue
    Green
    cyan
    Red
    Magenta
    Yellow
    White
End Enum

Public CRLF As String ' Cairrage return/Line feed
Public OldText As String ' Holds any text still
                      ' needing processed
Public channel As String ' Holds the channel name
Public Nickname As String 'holds the nick name
Public username As String 'holds the real name
Public idport As Long 'holds the identd port
Public userid As String 'holds the userid for identd
Public ostype As String 'holds the ostype for the identd
Public CMode  As Integer ' CurrentMode of client
Public text$ 'converted thana text
Public Language  As Languages
Public ConvertOnot As Boolean 'holds if the msg to be converted to thaana or not, Ture is conver  False is No
Public server As String 'server hostname
Public port As Long 'port number
Public buffer As String

Public servers() As serverType 'stores all the servers and info about it = simon
Public chanState()  As whichChannel           ' Array of channel names you we are in
Public chanCount As Integer 'number of channels open
Public channels() As New channelF   ' Array of child form objects

Public PvtWindows() As New PvtForm 'for the pvt windows collection --Echo
Public PvtWindowState() As whichChannel 'array of private windows  --Echo

Sub ChangePvtWindowNick(oldNicK As String, NewNicK As String)
'change ownership of the oldnick 's windows to NewNicK s....
'to be xecuted when a user changes nick ---Echo ..
    Dim i As Integer
    Dim ArrayCount As Integer
    
    ArrayCount = UBound(PvtWindows)

    For i = 1 To ArrayCount
        If Not PvtWindowState(i).deleted Then 'do not include deleted windows
            If PvtWindows(i).Caption = oldNicK Then
                PvtWindows(i).Caption = NewNicK
                PvtWindows(i).Tag = Str(i)
                PvtWindowState(i).name = NewNicK
                Exit For
            End If
        End If
    Next

End Sub


Function FreePvtIndex() As Integer
'-- find the free index for the private window
'----Echo
On Error GoTo Errhandler

    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(PvtWindows)

    ' Cycle through the Pvtwindows array. If one of the
    ' window has been deleted, then return that index.
    For i = 1 To ArrayCount
        If PvtWindowState(i).deleted Then
            FreePvtIndex = i
            PvtWindowState(i).deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the pvt window array have
    ' been deleted, then increment the window and the
    ' pvtwindowstate arrays by one and return the index to the
    ' new element.
    ReDim Preserve PvtWindows(ArrayCount + 1)
    ReDim Preserve PvtWindowState(ArrayCount + 1)
    FreePvtIndex = UBound(PvtWindows)

Exit Function
Errhandler:
If VBA.Information.err.Number = 9 Then 'subscript out of bound, gonna occur for 1st window
    Resume Next 'just resume.. not an actual err
End If
Resume Next
End Function

'get n return the index of window,,, using nick
Function GetPvtIndex(nick As String) As Integer

    Dim i As Integer
    Dim ArrayCount As Integer
    
    ArrayCount = UBound(PvtWindows)
    For i = 1 To ArrayCount
        If Not PvtWindowState(i).deleted Then 'do not include deleted windows
            If PvtWindows(i).Caption = nick Then
                GetPvtIndex = i
                Exit For
            End If
        End If
    Next
   

End Function

'write the msg to pvt window
'---echo
Sub WriteToPvtWindow(nick As String, msg As String)
    Dim PvtIndex As Integer
    
    If Not DoesPvtExists(nick) Then 'if no pvt for sender
        CreatePvtWindow (nick)          'open a new one
    End If
    PvtIndex = GetPvtIndex(nick)
    PvtWindows(PvtIndex).WindowState = vbNormal
    PvtWindows(PvtIndex).AddText "<" & nick & ">" & Mid(msg, InStr(msg, ":") + 1, Len(msg) - InStr(msg, ":")), QBColor(0) 'with black color
                

End Sub

'Creates a new Private window
'------- Echo
Sub CreatePvtWindow(nick As String)
Dim PvtIndex As Integer

PvtIndex = FreePvtIndex
PvtWindows(PvtIndex).Tag = PvtIndex 'the tag will contain the array index of the window
PvtWindows(PvtIndex).Caption = nick 'change the caption to nick
PvtWindows(PvtIndex).Show 'display the window
PvtWindowState(PvtIndex).name = nick 'just

'set the language of window from the default language
If Language = lngEnglish Then
    PvtWindowState(PvtIndex).Lang = lngEnglish
    Call ClientMain.mnuLangEnglish_Click 'this is needed explicity cox when a form is created it won't fire the Activate event
Else
    PvtWindowState(PvtIndex).Lang = lngThaana
    ClientMain.mnuLangThaana_Click 'this is needed explicity cox when a form is created it won't fire the Activate event
End If
    
End Sub

'checks where a window exists for a particulat nick
'----Echo
Function DoesPvtExists(nick As String) As Boolean
On Error GoTo Errhandler
    Dim i As Integer
    Dim ArrayCount As Integer
    
    DoesPvtExists = False    'initialize
    ArrayCount = UBound(PvtWindows) 'if no windows have been created this will occur an error..
                                    'so the err is trapped..

    For i = 1 To ArrayCount
        If Not PvtWindowState(i).deleted Then 'do not include deleted windows
            If PvtWindows(i).Caption = nick Then
                DoesPvtExists = True
                Exit For
            End If
        End If
    Next
    
Exit Function
Errhandler:
If VBA.Information.err = 9 Then
    Resume Next
End If

End Function
Sub newChannel()
    Dim fIndex As Integer
    
    ' Find the next available index and show the child form.
    fIndex = FindFreeIndex()
    chanState(fIndex).name = channel
    chanState(fIndex).deleted = False ' cos we just joined it man
    chanState(fIndex).Lang = lngEnglish 'change to default lang
    channels(fIndex).Tag = fIndex
    channels(fIndex).Caption = channel
    channels(fIndex).setChanName channel
    channels(fIndex).Show
   
End Sub

Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(channels)
    chanCount = ArrayCount + 1
    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    'For i = 1 To ArrayCount
    '    If FState(i).Deleted Then
    '        FindFreeIndex = i
    '        FState(i).Deleted = False
    '        Exit Function
    '    End If
    'Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve channels(chanCount)
    ReDim Preserve chanState(chanCount)
    FindFreeIndex = UBound(channels)
    'MsgBox FindFreeIndex
    
End Function


'this function checks to see if the given command text is
'really the command in the string. - simon
'note: just trying to make code more readable
Function this_command(text As String, this As String) As Boolean
    If this = Mid$(text, InStr(text, " ") + 1, Len(this)) Then
        this_command = True
    Else
        this_command = False
    End If
End Function

'this function gets the nickname of the person from the
'text that the server sends - simon
Function getNick(text As String) As String
    getNick = Mid$(text, 2, InStr(text, "!") - 2)
End Function

Function getChanIndex(chanName As String) As Integer
    Dim count As Integer
    Dim i As Integer
    count = UBound(chanState)
    For i = 0 To count
        If chanState(i).name = chanName Then
            getChanIndex = i
        End If
    Next
    
End Function
'for reading in our personal data
'and assigning them to the variables
Sub readUserINI()
    Dim rec As String
    Dim conf As String
    Dim data As String
    Open App.Path + "/user.ini" For Input As #1
    Input #1, rec 'read off the first line
    Do While Not EOF(1) ' Loop until end of file.
        Input #1, rec    ' Read data into two variables.
        conf = Left$(rec, (InStr(1, rec, "=") - 1))
        
        data = Trim(Right$(rec, Len(rec) - (Len(conf) + 1)))
        If conf = "server" Then
            
            server = data
        ElseIf conf = "realname" Then
            username = data
        ElseIf conf = "nickname" Then
            Nickname = data
        ElseIf conf = "userid" Then
            userid = data
        ElseIf conf = "ostype" Then
            ostype = data
        ElseIf conf = "port" Then
            idport = CLng(data)
        End If
        
        
        'Debug.Print conf & " > " & data  ' Print data to Debug window.
    Loop
    Close #1
End Sub

Sub readServerList()
    Dim rec As String 'holds the complete string from file
    
    Dim data As String 'holds server name from file
    Dim sIP As String 'holds server IP from file
    Dim index As Integer
    index = 0

    Dim count, start, ending As Integer
    count = 1
    
    Open App.Path + "\servers.ini" For Input As #1
    Do While Not EOF(1)
        Line Input #1, rec
        data = Trim(Mid$(rec, 1, InStr(1, rec, "SERVER") - 1))
        setup.ServerCombo.AddItem data, index
        start = InStr(1, rec, ":") + 1
        start = InStr(start, rec, ":") + 1
        ending = InStr((start), rec, ":")
        sIP = Mid$(rec, start, (ending - start))
        ReDim Preserve servers(count)
        With servers(count)
            .IP = sIP
            .name = data
            .port = 6667
        End With
        count = count + 1
        index = index + 1
        'Debug.Print data
    Loop
End Sub

'function used to generate RGB colour codes for the
'server window and the channel windows - simon

Function setcolor(kula As colour) As Long

    If kula = Black Then
        setcolor = RGB(0, 0, 0)
    ElseIf kula = Blue Then
        setcolor = RGB(24, 54, 154)
    ElseIf kula = Green Then
        setcolor = RGB(0, 128, 0)
    ElseIf kula = cyan Then
        setcolor = RGB(0, 255, 255)
    ElseIf kula = Red Then
        setcolor = RGB(198, 0, 0)
    ElseIf kula = Magenta Then
        setcolor = RGB(168, 19, 168)
    ElseIf kula = Yellow Then
        setcolor = RGB(192, 141, 22)
    ElseIf kula = White Then
        setcolor = RGB(255, 255, 255)
    End If
End Function

Sub ExecuteAction(Action As String, msg As String, CallForm As Object)
      Select Case Action
        Case "JOIN"
            ClientMain.SendData "JOIN " & msg: CMode = 1  ' join the channel, set the mode
            channel = Trim(msg)
            newChannel
        Case "ME"
          ' if we're in a channel, then do an action
          'If Namelist.ListCount > 0 Then ClientMain.SendData "PRIVMSG " & channel & " :" & Chr$(1) & "ACTION " & msg & Chr$(1)
          CallForm.AddText "* " & Nickname & " " & msg
        Case "MSG"
          ' send a priv msg
          ClientMain.SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " :" & Mid$(msg, InStr(msg, " ") + 1)
          ConvertOnot = True
          CallForm.AddText "=->" & Left$(msg, InStr(msg, " ") - 1) & "<-= " & Mid$(msg, InStr(msg, " ") + 1)
        Case "VERSION"
            ClientMain.SendData "PRIVMSG " & Left$(msg, InStr(msg, " ") - 1) & " : VERSION"
            ConvertOnot = False
            CallForm.AddText "[VERSION] " & Left$(msg, InStr(msg, " ") - 1)
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

End Sub


