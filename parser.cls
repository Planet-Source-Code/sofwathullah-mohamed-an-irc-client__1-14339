VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class performs the parsing of IRC server messages for easy retreival."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'>>>:vey!vey@pc140.pbs.port.ac.uk PRIVMSG #kotari :need to call some one n get some info....
'>>>:kandu!ultra@infinite1-27.inf.net.au PRIVMSG #kotari :ehentha?
'>>>:rxs!rxs@Flex079.GIRS.WAU.NL PRIVMSG #kotari :ACTION is from henveiru Baburuvaadhe!
'>>>:kandu!ultra@infinite1-27.inf.net.au PRIVMSG #kotari :ok than babe
'>>>:rxs!rxs@Flex079.GIRS.WAU.NL PRIVMSG #kotari :hehehe
'>>>:unit!unit@mel-as2-dial-8.vic.hotkey.net.au PART #kotari

'>>>:kandu!ultra@infinite2-104.inf.net.au PRIVMSG #kotari :haadha slow eh!!!!!!!!!!
'>>>:lessie!~user@as1_async13.dhivehinet.net.mv PRIVMSG #kotari :untitled
'>>>:kandu!ultra@infinite2-104.inf.net.au PRIVMSG #kotari :alberta
'>>>:kandu!ultra@infinite2-104.inf.net.au PRIVMSG #kotari :alberta
'>>>:ohuioj!fgffg@dialup-melax563.mpx.com.au JOIN :#kotari
'>>>:kandu!ultra@infinite2-104.inf.net.au PRIVMSG #kotari :where you been so long
'>>>:BrAiNdEaD!judas_prie@203.143.0.42 QUIT :Dallas-R.Tx.US.Undernet.org Vancouver.BC.CA.Undernet.Org
'>>>:kandu!ultra@infinite2-104.inf.net.au PRIVMSG #kotari :eheheh

'local variable(s) to hold property value(s)
Private mvarusernick As String 'local copy
Private mvaruserhost As String 'local copy
Private mvarmessage As String 'local copy
'local variable(s) to hold property value(s)
Private mvarmsgType As String 'local copy
'local variable(s) to hold property value(s)
Private mvarchannelName As String 'local copy
Public Property Let channelName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.channelName = 5
    mvarchannelName = vData
End Property


Public Property Get channelName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.channelName
    channelName = mvarchannelName
End Property



Public Property Let msgType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.msgType = 5
    mvarmsgType = vData
End Property


Public Property Get msgType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.msgType
    msgType = mvarmsgType
End Property

Public Sub parse(text As String)
    Exit Sub 'temp out of service :)
    Dim temp As String
    Dim unick As String 'temp storage for nickname
    Dim uhost As String 'temp storage for hostname
    Dim mType As String 'temp storage for message type
    Dim start As Integer
    Dim cName As String 'temp storage for channel name
    temp = Trim(text)
    start = 2 'lets ignore the ":" at the start of mesg
    unick = Mid$(temp, start, (InStr(temp, "!") - 2)) 'get the nickname
    start = Len(unick) + 3 'change start to end of nick and "!"
    uhost = Trim(Mid$(temp, start, (InStr(1, temp, " ") - start)))
    start = start + Len(uhost) + 1 'set the lenght  to what it says
    mType = Trim(Mid$(temp, start, InStr(start, temp, " ") - start))
    start = start + Len(mType) + 2 'increment the start again
    cName = Trim(Mid$(temp, start, start)) ' don't know ehy but gives an error normal way, so i removed th instr part -sofwath
    start = start + Len(cName) + 3
    'mvarmessage = Trim(Mid$(temp, start, (Len(temp) - (start - 1))))
    mvarusernick = unick
    mvaruserhost = uhost
    mvarmsgType = mType
    mvarchannelName = cName
End Sub

Public Property Let message(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.message = 5
    mvarmessage = vData
End Property


Public Property Get message() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.message
    message = mvarmessage
End Property



Public Property Let userhost(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.userhost = 5
    mvaruserhost = vData
End Property


Public Property Get userhost() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.userhost
    userhost = mvaruserhost
End Property



Public Property Let usernick(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.usernick = 5
    mvarusernick = vData
End Property


Public Property Get usernick() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.usernick
    usernick = mvarusernick
End Property



