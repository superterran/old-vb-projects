Attribute VB_Name = "Module1"
'put this in a module
'i did not create this code. it was made by revolt (www.revoltsoft.com)
'i got it from www.gamesxposed.com
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

'put this in the same module
Public Sub SaveKey(hkey As Long, strPath As String)
Dim r
    Dim keyhand&
    r = RegCreateKey(hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(hkey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
Dim r
Dim lValueType
    r = RegOpenKey(hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
    If GetString = "" Then
    GetString = ""
    End If
End Function


Public Function SaveString(hkey As Long, strPath As String, strValue As String, strData As String)
Dim keyhand As Long
Dim r As Long
    r = RegCreateKey(hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    r = RegCloseKey(keyhand)
    If r = 0 Then
    SaveString = ""
    Else
    SaveString = ""
    End If
End Function


Function GetDWord(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
    If GetDWord = "" Then
    GetDWord = ""
    End If
End Function

Function SaveDword(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
Dim lResult As Long
Dim keyhand As Long
Dim r As Long
    r = RegCreateKey(hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
    If r = 0 Then
    SaveDword = ""
    Else
    SaveDword = ""
    End If
End Function

Public Function DeleteKey(ByVal hkey As Long, ByVal strKey As String)
Dim r As Long
    r = RegDeleteKey(hkey, strKey)
    If r = 0 Then
    DeleteKey = ""
    Else
    DeleteKey = ""
    End If
End Function

Public Function DeleteValue(ByVal hkey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
Dim r
    r = RegOpenKey(hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
    If r = 0 Then
    DeleteValue = ""
    Else
    DeleteValue = ""
    End If
End Function


