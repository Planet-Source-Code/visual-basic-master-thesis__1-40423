VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
' Programmer  : Asian College OF Technology
' Source code : Registry Class
'******************************************************************

' Variable Declaration
Dim OHandle As Long

' constant declaration
Option Explicit
Const REG_DWORD = 4                      ' 32-bit number
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_NONE = 0

' type declaration
Public Enum Hive
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum KeyType
    KSTRING = REG_SZ
    KDWORD = REG_DWORD
End Enum

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

' api function's declarations
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Property Get RegKeyOpenInstance() As Long
    RegKeyOpenInstance = OHandle
End Property

Public Function CreateKey(ByVal Mkey As Hive, ByVal SubKey As String) As Long
    RegCreateKey Mkey, SubKey, OHandle      ' invoke api function
    CreateKey = OHandle                     ' return variable
End Function

Public Function Openkey(ByVal Mkey As Hive, ByVal SubKey As String) As Long
    RegOpenKey Mkey, SubKey, OHandle
    Openkey = OHandle
End Function

Public Sub DeleteKey(ByVal Mkey As Hive, ByVal SubKey As String)
    RegDeleteKey Mkey, SubKey
End Sub

Public Sub SetValue(ByVal Openkey As Long, ByVal typ As KeyType, ByVal ValueName As String, data As Variant)
    Dim m_svalue As String
    Dim m_lvalue As Long
    If typ = KSTRING Then
        m_svalue = data
        RegSetValueEx Openkey, ValueName, 0&, typ, ByVal m_svalue, Len(m_svalue)
    ElseIf typ = KDWORD Then
        m_lvalue = data
        RegSetValueEx Openkey, ValueName, 0&, typ, m_lvalue, 4
    End If
End Sub

Public Function RetrieveData(ByVal Openkey As Long, ByVal typ As KeyType, ByVal ValueName As String)
    If typ = KSTRING Then
        Dim sBuffer As String * 255
        sBuffer = String(255, " ")
        RegQueryValueEx Openkey, ValueName, 0, REG_SZ, ByVal sBuffer, Len(sBuffer)
        RetrieveData = Trim(sBuffer)
    ElseIf typ = KDWORD Then
        Dim dBuffer As Long
        RegQueryValueEx Openkey, ValueName, 0, REG_DWORD, dBuffer, Len(dBuffer)
        RetrieveData = dBuffer
    End If
End Function

Public Sub DeleteRegVal(ByVal Openkey As Long, ByVal ValueName As String)
    RegDeleteValue Openkey, ValueName
End Sub

Public Sub CloseRegistry(ByVal Key As Long)
    RegCloseKey Key
End Sub
