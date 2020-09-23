Attribute VB_Name = "APIdefs"
' to read the ini files
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'to write the ini files
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'--------------------------------------------------------
'Win32 API calls
'--------------------------------------------------------

Function writeUserINI(ByVal section As String, ByVal key As String, ByVal value As String)
    Dim s&
    Dim filename As String
    filename = App.Path & "/user.ini"
    s& = WritePrivateProfileString(ByVal section, ByVal key, ByVal value, ByVal filename)
    writeUserINI = 0
End Function
