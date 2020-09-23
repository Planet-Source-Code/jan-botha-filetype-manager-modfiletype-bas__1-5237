Attribute VB_Name = "RegistryAccess"
' Win32 Registry Access Module
'
' WINREG32.BAS - Copyright <C> 1998, 1999 Randy Mcdowell.
'
' If you modify this code please send me a copy, it's not commented
' really well so you'll have to bear with me here. I have included some
' sample subroutines and  functions to  access the registry. I have  a
' more complex  module  much  more  rich in  code if you want it you
' will need to Email me and ask for it.  mcdowellrandy@hotmail.com.

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type

Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Const ERROR_SUCCESS = 0&
Public Const REG_BINARY = 3                                        ' Free form binary
Public Const REG_CREATED_NEW_KEY = &H1               ' New Registry Key created
Public Const REG_DWORD = 4                                        ' 32-bit number
Public Const REG_DWORD_BIG_ENDIAN = 5                   ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4               ' 32-bit number (same as REG_DWORD)
Public Const REG_EXPAND_SZ = 2                                  ' Unicode nul terminated string
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_LINK = 6                                               ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                                       ' Multiple Unicode strings
Public Const REG_NONE = 0                                             ' No value type
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4     ' Time stamp
Public Const REG_NOTIFY_CHANGE_NAME = &H1            ' Create or delete (child)
Public Const REG_OPENED_EXISTING_KEY = &H2            ' Existing Key opened
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8
Public Const REG_OPTION_BACKUP_RESTORE = 4          ' Open for backup or restore
Public Const REG_OPTION_CREATE_LINK = 2                   ' Created key is a symbolic link
Public Const REG_OPTION_NON_VOLATILE = 0                 ' Key is preserved when system is rebooted
Public Const REG_OPTION_RESERVED = 0                       ' Parameter is reserved
Public Const REG_OPTION_VOLATILE = 1                          ' Key is not preserved when system is rebooted
Public Const REG_REFRESH_HIVE = &H2                          ' Unwind changes to last flush
Public Const REG_RESOURCE_LIST = 8                             ' Resource list in the resource map
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Public Const REG_SZ = 1                                                   ' Unicode nul terminated string
Public Const REG_WHOLE_HIVE_VOLATILE = &H1            ' Restore whole hive volatile
Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
Public Sub CreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant)

    Dim hHnd As Long
    
    If Not IsMissing(SubKey) Then
        Temp = RegCreateKey(hKey, Key & "\" & SubKey, hHnd)
        Temp = RegCloseKey(hHnd)
    Else
        Temp = RegCreateKey(hKey, Key, hHnd)
        Temp = RegCloseKey(hHnd)
    End If

End Sub

Public Function GetString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String)

    Dim hHnd As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lValueType As Long
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufferSize, " ")
        lResult = RegQueryValueEx(hHnd, ValueName, 0&, 0&, ByVal strBuf, lDataBufferSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        
        End If
    End If

End Function

Public Sub SaveString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As String)

    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    Temp = RegSetValueEx(hHnd, ValueTitle, 0, REG_SZ, ByVal ValueData, Len(ValueData))
    Temp = RegCloseKey(hHnd)

End Sub



Public Function GetDWord(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As Long

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufferSize As Long
    Dim Temp As Long
    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lDataBufferSize = 4
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, lBuf, lDataBufferSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If

    Temp = RegCloseKey(hHnd)

End Function

Public Function GetBinary(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As Long

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufferSize As Long
    Dim Temp As Long
    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lDataBufferSize = 4
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, lBuf, lDataBufferSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            GetBinary = lBuf
        End If
    End If

    Temp = RegCloseKey(hHnd)

End Function


Public Sub SaveDWord(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As Long)

    Dim lResult As Long
    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    lResult = RegSetValueEx(hHnd, ValueTitle, 0&, REG_DWORD, ValueData, 4)
    Temp = RegCloseKey(hHnd)

End Sub
Public Sub SaveBinary(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As Long)

    Dim lResult As Long
    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    lResult = RegSetValueEx(hHnd, ValueTitle, 0&, REG_BINARY, ValueData, 4)
    Temp = RegCloseKey(hHnd)

End Sub




Public Sub DeleteKey(ByVal hKey As Long, ByVal Key As String)

    Dim Temp As Long
    
    Temp = RegDeleteKey(hKey, Key)

End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal Value As String)

    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    Temp = RegDeleteValue(hHnd, Value)
    Temp = RegCloseKey(hHnd)

End Sub

