Attribute VB_Name = "WinReg32"
'
' Win32 Registry Access Module (Third Version)
'
' ~~ Enriched using Microsoft Platform SDK ~~~
'
' WINREG32.BAS - Copyright <C> 1998, 2000 Randy Mcdowell.
'
' In the last release I discovered a few errors, so I took
' my time on this  one  to  re-write every little piece of
' code, using my old  module as a  building block. In this
' version I added the 'OS' suffix to all  the registry API
' declarations so I could use 'Reg' in my  function names.
' Now  all  Subs  are  Functions, and  return some type of
' value, depending  on  the  procedure. In this  version I
' also  added  the  ability to save a  string  to a binary
' key (RegWriteExtended), which was requested. Also  error
' checking has been  revised and the RegLastError variable
' always contains the  last  registry  error that occured.
' If you modify this code  please send  me a copy so I can
' learn from what  you  have  done. I  hold  absolutely no
' warranties on this code  and  I am not  responsible  for
' any damage it does to your  registry  or  your computer.

Option Explicit

' Temporary Stack Storage Variables
Private Temp&, TempEx&, TempExA$
Private TempExB&, TempExC%

' Handle And Other Storage Variables
Private hHnd&, lpAttr As SECURITY_ATTRIBUTES
Private KeyPath$, hDepth&

' Variable To Hold Last Error
Public RegLastError As Long

' Reg Basic API Functions
Declare Function OSRegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function OSRegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function OSRegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function OSRegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function OSRegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function OSRegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Declare Function OSRegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function OSRegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Declare Function OSRegFlushKey Lib "advapi32.dll" Alias "RegFlushKey" (ByVal hKey As Long) As Long
Declare Function OSRegGetKeySecurity Lib "advapi32.dll" Alias "RegGetKeySecurity" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long
Declare Function OSRegSetKeySecurity Lib "advapi32.dll" Alias "RegSetKeySecurity" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function OSRegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Declare Function OSRegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function OSRegNotifyChangeKeyValue Lib "advapi32.dll" Alias "RegNotifyChangeKeyValue" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Declare Function OSRegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function OSRegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function OSRegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function OSRegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

' Reg Extended API Functions
Declare Function OSRegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function OSRegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function OSRegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function OSRegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function OSRegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function OSRegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function OSRegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function OSRegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function OSRegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

' Reg Return Error Constants
Public Const ERROR_SUCCESS = 0&                  ' Operation Successfull
Public Const ERROR_NONE = 0                      ' No Errors
Public Const ERROR_BADDB = 1                     ' Corrupt Registry Database
Public Const ERROR_BADKEY = 2                    ' Key Name Is Bad
Public Const ERROR_CANTOPEN = 3                  ' Unable To Open Key
Public Const ERROR_CANTREAD = 4                  ' Unable To Read Key
Public Const ERROR_CANTWRITE = 5                 ' Unable To Write Key
Public Const ERROR_OUTOFMEMORY = 6               ' Out Of Memory
Public Const ERROR_ARENA_TRASHED = 7             ' Unknown Error
Public Const ERROR_ACCESS_DENIED = 8             ' Registry Access Denied
Public Const ERROR_INVALID_PARAMETERS = 87       ' Invalid Parameter
Public Const ERROR_NO_MORE_ITEMS = 259           ' No More Items

' Reg Key ROOT Locations
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

' I Thought You Would Like This
Public Const HKCR = HKEY_CLASSES_ROOT
Public Const HKCC = HKEY_CURRENT_CONFIG
Public Const HKCU = HKEY_CURRENT_USER
Public Const HKDD = HKEY_DYN_DATA
Public Const HKLM = HKEY_LOCAL_MACHINE
Public Const HKPD = HKEY_PERFORMANCE_DATA
Public Const HKUS = HKEY_USERS

' Reg Value Data Types
Public Const REG_NONE = 0                        ' No value type
Public Const REG_SZ = 1                          ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                   ' Unicode nul terminated string
Public Const REG_BINARY = 3                      ' Free form binary
Public Const REG_DWORD = 4                       ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4         ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5            ' 32-bit number
Public Const REG_LINK = 6                        ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                    ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8               ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9    ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10 '
Public Const REG_CREATED_NEW_KEY = &H1           ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY = &H2       ' Existing Key opened
Public Const REG_WHOLE_HIVE_VOLATILE = &H1       ' Restore whole hive volatile
Public Const REG_REFRESH_HIVE = &H2              ' Unwind changes to last flush
Public Const REG_NOTIFY_CHANGE_NAME = &H1        ' Create or delete (child)
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2  '
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4    ' Time stamp
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8    '

' Reg Create Type Values
Public Const REG_OPTION_RESERVED = 0             ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE = 0         ' Key is preserved when system is rebooted
Public Const REG_OPTION_VOLATILE = 1             ' Key is not preserved when system is rebooted
Public Const REG_OPTION_CREATE_LINK = 2          ' Created key is a symbolic link
Public Const REG_OPTION_BACKUP_RESTORE = 4       ' Open for backup or restore

' Reg Legal Options (Whats This?)
Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

' Reg Key Security Options
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum

' Reg API Type Structures
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

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Function RegCheckError(ByRef ErrorValue As Long) As Boolean

        If ((ErrorValue < 8) And (ErrorValue > 1)) Or _
        (ErrorValue = 87) Or (ErrorValue = 259) Then _
        RegCheckError = -1 Else RegCheckError = 0

End Function


Public Function RegConnectRegistry(ByRef Computer As String) As Long
    
        ' Connect To The Network Computer
        Temp& = OSRegConnectRegistry(Computer, 0&, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ConnectRegError
    
        ' Operation Was Successful
        RegConnectRegistry = hHnd&

        ' Exit Function With Passed Value
        Exit Function

ConnectRegError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegConnectRegistry = 0
    
End Function
Public Function RegDeleteValue(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal Value As String) As Boolean
    
        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Open The Key For Operations
        Temp& = OSRegOpenKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo DeleteValueError
    
        ' Delete Existing Value From Key
        Temp& = OSRegDeleteValue(hHnd&, Value)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo DeleteValueError
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
        ' Operation Was Successful
        RegDeleteValue = -1

        ' Exit Function With Passed Value
        Exit Function

DeleteValueError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegDeleteValue = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegFlushKey(ByVal Key As Long) As Boolean
    
        ' Flush The Specified Key
        Temp& = OSRegFlushKey(Key)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo FlushKeyError
    
        ' Operation Was Successful
        RegFlushKey = -1

        ' Exit Function With Passed Value
        Exit Function

FlushKeyError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegFlushKey = 0
    
End Function
Public Function RegReadBinary(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As Long

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Open The Key For Operations
        Temp& = OSRegOpenKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadBinaryError
    
        ' Read In Information In Binary Format
        Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, TempEx&, Temp&, 4&)
    
        ' Operation Was Successful
        If TempEx& = REG_BINARY Then RegReadBinary = Temp&

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Exit Function With Passed Value
        Exit Function

ReadBinaryError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegReadBinary = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function

Public Function RegReadDWord(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As Long

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Open The Key For Operations
        Temp& = OSRegOpenKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadDWordError
    
        ' Read In Information In Binary Format
        TempExB& = OSRegQueryValueEx(hHnd&, ValueName, 0&, TempEx&, Temp&, 4&)

        ' Operation Was Successful
        If TempEx& = REG_DWORD Then RegReadDWord = Temp&

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Exit Function With Passed Value
        Exit Function

ReadDWordError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegReadDWord = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegWriteBinary(ByVal hKey As Long, _
        ByVal Key As String, _
        ByVal SubKey As String, _
        ByVal ValueName As String, _
        ByVal Value As Variant) As Boolean

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Create Key If It Doesn't Exist
        Temp& = OSRegCreateKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteBinaryError
    
        ' Set New Value For The Opened Key
        Temp& = OSRegSetValueEx(hHnd&, ValueName, 0&, REG_BINARY, Value, 4&)
     
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteBinaryError

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
        ' Operation Was Successful
        RegWriteBinary = -1

        ' Exit Function With Passed Value
        Exit Function

WriteBinaryError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegWriteBinary = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegWriteString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String, ByVal Value As String) As Boolean

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Create Key If It Doesn't Exist
        Temp& = OSRegCreateKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteStringError
    
        ' Set New Value For The Opened Key
        Temp& = OSRegSetValueEx(hHnd&, ValueName, 0&, REG_SZ, ByVal Value, Len(Value))
     
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteStringError

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Operation Was Successful
        RegWriteString = -1

        ' Exit Function With Passed Value
        Exit Function

WriteStringError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegWriteString = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegWriteExtended(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String, ByVal Value As String) As Boolean

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Create Key If It Doesn't Exist
        Temp& = OSRegCreateKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteExtendedError
    
        ' Set New Value For The Opened Key
        Temp& = OSRegSetValueEx(hHnd&, ValueName, 0&, REG_EXPAND_SZ, ByVal Value, Len(Value))
     
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteExtendedError

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Operation Was Successful
        RegWriteExtended = -1

        ' Exit Function With Passed Value
        Exit Function

WriteExtendedError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegWriteExtended = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function

Public Function RegWriteDWord(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String, ByVal Value As Long) As Boolean

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Create Key If It Doesn't Exist
        Temp& = OSRegCreateKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteDWordError
    
        ' Set New Value For The Opened Key
        Temp& = OSRegSetValueEx(hHnd&, ValueName, 0&, REG_DWORD, Value, 4&)
     
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo WriteDWordError

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Operation Was Successful
        RegWriteDWord = -1

        ' Exit Function With Passed Value
        Exit Function

WriteDWordError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegWriteDWord = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function


Public Function RegReadString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As String
        On Error Resume Next
        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Open The Key For Operations
        Temp& = OSRegOpenKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadStringError
    
        ' Read In Information In Unicode Format
        Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, TempEx&, Temp&, TempExB&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadStringError
    
        ' Operation Was Successful
        If TempEx& = REG_SZ Then
         
            ' Create ASCIIZ Based String
            TempExA$ = String(TempExB&, " ")
        
            ' Convert Information To String Format
            Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, 0&, ByVal TempExA$, TempExB&)

            ' Process Returned Information
            If RegCheckError(Temp&) Then GoTo ReadStringError
        
            ' Find Unicode String NULL Terminator
            TempExC% = InStr(TempExA$, Chr$(0))
        
            ' Return All Characters Before NULL
            If TempExC% > 0 Then
                RegReadString = Left$(TempExA$, TempExC% - 1)
            Else
                RegReadString = TempExA$
            End If

        End If

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Exit Function With Passed Value
        Exit Function

ReadStringError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegReadString = vbNullString
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegReadExtended(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As String

        ' Combine The Key And SubKey Paths
        If Not SubKey = "" Then KeyPath$ = _
        Key + "\" + SubKey Else KeyPath$ = Key
    
        ' Open The Key For Operations
        Temp& = OSRegOpenKey(hKey, KeyPath$, hHnd&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadExtendedError
    
        ' Read In Information In Unicode Format
        Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, TempEx&, Temp&, TempExB&)
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadExtendedError
    
        ' Operation Was Successful
        If TempEx& = REG_EXPAND_SZ Then
         
            ' Create ASCIIZ Based String
            TempExA$ = String(TempExB&, " ")
        
            ' Convert Information To String Format
            Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, 0&, ByVal TempExA$, TempExB&)

            ' Process Returned Information
            If RegCheckError(Temp&) Then GoTo ReadExtendedError
        
            ' Find Unicode String NULL Terminator
            TempExC% = InStr(TempExA$, Chr$(0))
        
            ' Return All Characters Before NULL
            If TempExC% > 0 Then
                RegReadExtended = Left$(TempExA$, TempExC% - 1)
            Else
                RegReadExtended = TempExA$
            End If

        End If

        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)

        ' Exit Function With Passed Value
        Exit Function

ReadExtendedError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegReadExtended = vbNullString
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function

Public Function RegUpdateKey(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByRef Value As String) As Boolean
    
        ' Set Security Attributes To Defaults
        lpAttr.nLength = 50
        lpAttr.lpSecurityDescriptor = 0
        lpAttr.bInheritHandle = -1
    
        ' Create/Open Specified Key
        Temp& = OSRegCreateKeyEx(hKey, Key, 0, REG_SZ, REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, lpAttr, hHnd&, hDepth&)
                             
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo UpdateKeyError
    
        ' A Space Is Required For RegSetValueEx()
        If (Value = "") Then Value = Chr$(32)
    
        ' Set New Value For The Opened Key
        Temp& = OSRegSetValueEx(hHnd&, SubKey, 0, REG_SZ, Value, _
        LenB(StrConv(Value, vbFromUnicode)))
                       
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo UpdateKeyError
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
        ' Operation Was Successful
        RegUpdateKey = -1

        ' Exit Function With Passed Value
        Exit Function

UpdateKeyError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegUpdateKey = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegCreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant) As Boolean
    
        ' Create Key If It Doesn't Exist
        If Not IsMissing(SubKey) Then
            Temp& = OSRegCreateKey(hKey, Key & "\" & SubKey, hHnd&)
        Else
            Temp& = OSRegCreateKey(hKey, Key, hHnd&)
        End If
    
        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo CreateKeyError
   
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
        ' Operation Was Successful
        RegCreateKey = -1

        ' Exit Function With Passed Value
        Exit Function

CreateKeyError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegCreateKey = 0
    
        ' Close Handle To Key
        Temp& = OSRegCloseKey(hHnd&)
    
End Function
Public Function RegDeleteKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant) As Boolean
    
        ' Delete Existing Key
        If IsMissing(SubKey) Then
            Temp& = OSRegDeleteKey(hKey, Key)
        Else
            Temp& = OSRegDeleteKey(hKey, Key & "\" & SubKey)
        End If

        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo DeleteKeyError

        ' Operation Was Successful
        RegDeleteKey = -1

        ' Exit Function With Passed Value
        Exit Function

DeleteKeyError:
    
        ' Store Error In Variable
        RegLastError = Temp&
    
        ' Operation Was Not Successful
        RegDeleteKey = 0
    
End Function
