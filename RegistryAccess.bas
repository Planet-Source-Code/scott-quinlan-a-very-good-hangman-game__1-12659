Attribute VB_Name = "RegistryAccess"
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Const KEY_NOTIFY = &H10
Public Const ERROR_SUCCESS = 0&

Const REG_SZ = 1&
Const KEY_QUERY_VALUE = &H1&
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const DisplayErrorMsg = False

Dim MainKeyHandle As Long
Dim lBufferSize As Long
Dim rtn As Long
Dim hkey As Long
Dim sBuffer As String

Function GetStringValue(SubKey As String, Entry As String)

    Call ParseKey(SubKey, MainKeyHandle)

    If MainKeyHandle Then
        rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hkey)
        If rtn = ERROR_SUCCESS Then
            sBuffer = Space(255)
            lBufferSize = Len(sBuffer)
            rtn = RegQueryValueEx(hkey, Entry, 0, REG_SZ, sBuffer, lBufferSize)
            If rtn = ERROR_SUCCESS Then
                rtn = RegCloseKey(hkey)
                sBuffer = Trim(sBuffer)
                GetStringValue = Left(sBuffer, Len(sBuffer) - 1)
            Else
                GetStringValue = ""
                If DisplayErrorMsg = True Then
                    MsgBox ErrorMsg(rtn)
                End If
            End If
        Else
            GetStringValue = ""
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
    End If

End Function

Private Sub ParseKey(Keyname As String, keyhandle As Long)
    
    rtn = InStr(Keyname, "\")

    If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then
        If DisplayErrorMsg = True Then
            MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + Keyname
        End If
        Exit Sub
    ElseIf rtn = 0 Then
        keyhandle = GetMainKeyHandle(Keyname)
        Keyname = ""
    Else
        keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1))
        Keyname = Right(Keyname, Len(Keyname) - rtn)
    End If

End Sub

Function ErrorMsg(lErrorCode As Long) As String

    Select Case lErrorCode
        Case 1009, 1015
            GetErrorMsg = "The Registry Database is corrupt!"
        Case 2, 1010
            GetErrorMsg = "Bad Key Name"
        Case 1011
            GetErrorMsg = "Can't Open Key"
        Case 4, 1012
            GetErrorMsg = "Can't Read Key"
        Case 5
            GetErrorMsg = "Access to this key is denied"
        Case 1013
            GetErrorMsg = "Can't Write Key"
        Case 8, 14
            GetErrorMsg = "Out of memory"
        Case 87
            GetErrorMsg = "Invalid Parameter"
        Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
        Case Else
            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
    End Select

End Function

Function GetMainKeyHandle(MainKeyName As String) As Long

    Const HKEY_LOCAL_MACHINE = &H80000002

    Select Case MainKeyName
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
    End Select

End Function

Function SetStringValue(SubKey As String, Entry As String, Value As String)

    Call ParseKey(SubKey, MainKeyHandle)

    If MainKeyHandle Then
        rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hkey)
        If rtn = ERROR_SUCCESS Then
            rtn = RegSetValueEx(hkey, Entry, 0, REG_SZ, ByVal Value, Len(Value))
            If Not rtn = ERROR_SUCCESS Then
                If DisplayErrorMsg = True Then
                    MsgBox ErrorMsg(rtn)
                End If
            End If
            rtn = RegCloseKey(hkey)
        Else
            If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
            End If
        End If
    End If

End Function

Function CreateKey(SubKey As String)

    Call ParseKey(SubKey, MainKeyHandle)

    If MainKeyHandle Then
        rtn = RegCreateKey(MainKeyHandle, SubKey, hkey)
        If rtn = ERROR_SUCCESS Then
            rtn = RegCloseKey(hkey)
        End If
    End If

End Function


