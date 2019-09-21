Attribute VB_Name = "modRegistry"
Option Explicit

'declare constants
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const INFINITE = -1&

'declare external functions
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String

'declare variables
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

'check that data exists
If Not IsEmpty(Default) Then
    GetSettingString = Default
Else
    GetSettingString = ""
End If

'search registry and build string
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

    If lValueType = REG_SZ Then
        strBuffer = String(lDataBufferSize, " ")
        lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
        intZeroPos = InStr(strBuffer, Chr$(0))
        
        If intZeroPos > 0 Then
            GetSettingString = Left$(strBuffer, intZeroPos - 1)
        Else
            GetSettingString = strBuffer
        End If
    
    End If
    
Else

    GetSettingString = "Error!"

End If

lRegResult = RegCloseKey(hCurKey)

End Function
