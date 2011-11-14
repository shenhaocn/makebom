Attribute VB_Name = "mGenAPI"
'***************************************************************************************
'**ģ �� ����mRegKeys
'**˵    ����TP-LINK SMB Switch Product Line Hardware Group ��Ȩ����2011 - 2012(C)
'**�� �� �ˣ�Shenhao
'**��    �ڣ�2011-11-03 22:29:00
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����ע����дģ��
'**��    ����V3.6.10
'***************************************************************************************

Option Explicit

'==============================================================================
'-ע��� ��������...
'==============================================================================
'ע�������
Public Enum enumRegMainKey
   iHKEY_CURRENT_USER = &H80000001
   iHKEY_LOCAL_MACHINE = &H80000002
   iHKEY_CLASSES_ROOT = &H80000000
   iHKEY_CURRENT_CONFIG = &H80000005
   iHKEY_USERS = &H80000003
   iHKEY_PERFORMANCE_DATA = &H80000004
End Enum

'ע�����������
Public Enum enumRegSzType
   iREG_SZ = &H1                    ' �ַ���
   iREG_EXPAND_SZ = &H2             ' ��չ���������ַ���
   iREG_BINARY = &H3                ' ԭʼ�Ķ���������
   iREG_DWORD = &H4                 ' 32-bit ����
   iREG_NONE = 0&
   iREG_DWORD_LITTLE_ENDIAN = 4&
   iREG_DWORD_BIG_ENDIAN = 5&
   iREG_LINK = 6&
   iREG_MULTI_SZ = 7&
   iREG_RESOURCE_LIST = 8&
   iREG_FULL_RESOURCE_DEscriptOR = 9&
   iREG_RESOURCE_REQUIREMENTS_LIST = 10&
End Enum

'ע���ȫ��������...
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

' ע���������ֵ...
Const REG_OPTION_NON_VOLATILE = 0       ' ��ϵͳ��������ʱ���ؼ��ֱ�����

' ע��� �������
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1009&
Public Const ERROR_BADKEY = 1010&
Public Const ERROR_CANTOPEN = 1011&
Public Const ERROR_CANTREAD = 1012&
Public Const ERROR_CANTWRITE = 1013&
Public Const ERROR_OUTOFMEMORY = 14&
Public Const ERROR_INVALID_PARAMETER = 87&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_MORE_DATA = 234&

' ע���ؼ��ְ�ȫѡ��...
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const KEY_CREATE_LINK = &H20&

Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                               KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                               KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                               
                               
'==============================================================================
'-ע��� API ����...
'==============================================================================
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FileTime) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal lpSecurityAttributes As Long) As Long
Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwflags As Long) As Long


'==============================================================================
' ��ȡ�ļ�����ʱ�� �޸�ʱ�� ����ʱ�� ��������
'==============================================================================
Public Const OFS_MAXPATHNAME = 128
Public Const OF_READ = &H0
Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
End Type

Public Type FileTime
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type


Public Type BY_HANDLE_FILE_INFORMATION
     dwFileAttributes As Long
     ftCreationTime As FileTime
     ftLastAccessTime As FileTime
     ftLastWriteTime As FileTime
     dwVolumeSerialNumber As Long
     nFileSizeHigh As Long
     nFileSizeLow As Long
     nNumberOfLinks As Long
     nFileIndexHigh As Long
     nFileIndexLow As Long
End Type

Public Type TIME_ZONE_INFORMATION
     bias As Long
     StandardName(32) As Integer
     StandardDate As SYSTEMTIME
     StandardBias As Long
     DaylightName(32) As Integer
     DaylightDate As SYSTEMTIME
     DaylightBias As Long
End Type


'==============================================================================
' ��ȡ�ļ�����ʱ�� �޸�ʱ�� ����ʱ�� API����
'==============================================================================
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long
Public Const OF_READWRITE = &H2
 
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FileTime) As Long
Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'==============================================================================
' ������� ��������
'==============================================================================
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Public Const SWP_NOMOVE = &H2 '������Ŀǰ�Ӵ�λ��
Public Const SWP_NOSIZE = &H1 '������Ŀǰ�Ӵ���С
Public Const HWND_TOPMOST = -1 '�趨Ϊ���ϲ�
Public Const HWND_NOTOPMOST = -2 'ȡ�����ϲ��趨
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'==============================================================================
' ������� API��������
'==============================================================================
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


'==============================================================================
' ���� API��������
'==============================================================================
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'==============================================================================
' �������ָ��
'==============================================================================
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursorBynum& Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long)
Public Const IDC_HAND = 32649&
'*************************************************************************
'**�� �� ����GetValue
'**��    �룺ByVal mainKey(enumRegMainKey)       - ����
'**        ��ByVal subKey(String)                - �ӽ�
'**        ��ByVal keyV(String)                  - ����
'**        ��ByRef sValue(Variant)               - ��ֵ
'**        ��Optional ByRef rlngErrNum(Long)     - �����
'**        ��Optional ByRef rstrErrDescr(String) - ��������
'**��    ����(Boolean) -
'**������������ȡע����ֵ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-03 22:30:00
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.10
'*************************************************************************
Public Function GetValue(ByVal mainKey As enumRegMainKey, _
                         ByVal subKey As String, _
                         ByVal keyV As String, _
                         ByRef sValue As Variant, _
                         Optional ByRef rlngErrNum As Long, _
                         Optional ByRef rstrErrDescr As String) As Boolean
                        
   Dim hKey As Long, lType As Long, lBuffer As Long, sBuffer As String, lData As Long
   
   On Error GoTo GetValueErr
   
   GetValue = False
   
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
   End If
   
   If RegQueryValueEx(hKey, keyV, 0, lType, ByVal 0, lBuffer) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
   End If
   
   
   Select Case lType
      Case iREG_SZ
         lBuffer = 255
         sBuffer = Space(lBuffer)
         If RegQueryValueEx(hKey, keyV, 0, lType, ByVal sBuffer, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
         End If
         sValue = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
      Case iREG_EXPAND_SZ
         sBuffer = Space(lBuffer)
         If RegQueryValueEx(hKey, keyV, 0, lType, ByVal sBuffer, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
         End If
         sValue = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
      Case iREG_DWORD
         If RegQueryValueEx(hKey, keyV, 0, lType, lData, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
         End If
         sValue = lData
      Case iREG_BINARY
         If RegQueryValueEx(hKey, keyV, 0, lType, lData, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
         End If
         sValue = lData
   End Select
   
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
   End If
   
   GetValue = True
   
   Err.Clear

GetValueErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function

'����:ȡĳ��ֵ�µ�������
'����:RegEnumKeyVal
'����:hKey   RegMainKeyö��,subKey   �Ӽ�·������.
'����ֵ:String   �ַ�������
'����:
Public Function RegEnumKeyVal(ByVal mainKey As enumRegMainKey, _
                              ByVal subKey As String, _
                              ByRef sEValue() As String, _
                              Optional ByRef rlngErrNum As Long, _
                              Optional ByRef rstrErrDescr As String) As Boolean

    Dim hKey As Long, Idx As Long, sSave As String
    Dim RevVal() As String
    
    Dim strClassName     As String
    Dim lngClassNameLen     As Long
    Dim lngReserved     As Long
    Dim ftLast     As FileTime

    RegEnumKeyVal = False
    
    On Error GoTo RegEnumKeyValErr
    
    If RegOpenKeyEx(mainKey, subKey, 0, KEY_READ, hKey) <> ERROR_SUCCESS Then
       Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
    End If
    
    Idx = 0
    Do
        sSave = String(255, 0)
        If RegEnumKeyEx(hKey, Idx, sSave, 255, 0, strClassName, lngClassNameLen, ftLast) <> ERROR_SUCCESS Then Exit Do
        Idx = Idx + 1
        ReDim Preserve RevVal(Idx)
        RevVal(Idx - 1) = StripTerminator(sSave)
    Loop
    
    If RegCloseKey(hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "��ȡע���ֵʱ����"
    End If
    
    sEValue = RevVal
        
    RegEnumKeyVal = True
    
    Err.Clear
    
RegEnumKeyValErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
   
End Function


'*************************************************************************
'**�� �� ����SetValue
'**��    �룺ByVal mainKey(enumRegMainKey)       -
'**        ��ByVal subKey(String)                -
'**        ��ByVal keyV(String)                  -
'**        ��ByVal lType(enumRegSzType)          -
'**        ��ByVal sValue(Variant)               -
'**        ��Optional ByRef rlngErrNum(Long)     -
'**        ��Optional ByRef rstrErrDescr(String) -
'**��    ����(Boolean) -
'**��������������ע����ֵ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-03 22:31:06
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.10
'*************************************************************************
Public Function SetValue(ByVal mainKey As enumRegMainKey, _
                        ByVal subKey As String, _
                        ByVal keyV As String, _
                        ByVal lType As enumRegSzType, _
                        ByVal sValue As Variant, _
                        Optional ByRef rlngErrNum As Long, _
                        Optional ByRef rstrErrDescr As String) As Boolean
   Dim s As Long, lBuffer As Long, hKey As Long
   Dim ss As SECURITY_ATTRIBUTES
   
   On Error GoTo SetValueErr
   
   SetValue = False
   
   ss.nLength = Len(ss)
   ss.lpSecurityDescriptor = 0
   ss.bInheritHandle = True
   
   
   If RegCreateKeyEx(mainKey, subKey, 0, "", 0, KEY_WRITE, ss, hKey, s) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "����ע���ʱ����"
   End If
   
   
   Select Case lType
      Case iREG_SZ
         lBuffer = LenB(sValue)
         If RegSetValueEx(hKey, keyV, 0, lType, ByVal sValue, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "����ע���ʱ����"
         End If
      Case iREG_EXPAND_SZ
         lBuffer = LenB(sValue)
         If RegSetValueEx(hKey, keyV, 0, lType, ByVal sValue, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "����ע���ʱ����"
         End If
      Case iREG_DWORD
         lBuffer = 4
         If RegSetValueExA(hKey, keyV, 0, lType, sValue, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "����ע���ʱ����"
         End If
      Case iREG_BINARY
         lBuffer = 4
         If RegSetValueExA(hKey, keyV, 0, lType, sValue, lBuffer) <> ERROR_SUCCESS Then
            Err.Raise vbObjectError + 1, , "����ע���ʱ����"
         End If
      Case Else
         Err.Raise vbObjectError + 1, , "��֧�ָò�������"
   End Select
   
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "����ע���ʱ����"
   End If
   
   
   SetValue = True
   
   Err.Clear

SetValueErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function


'*************************************************************************
'**�� �� ����DeleteValue
'**��    �룺ByVal mainKey(enumRegMainKey)       -
'**        ��ByVal subKey(String)                -
'**        ��ByVal keyV(String)                  -
'**        ��Optional ByRef rlngErrNum(Long)     -
'**        ��Optional ByRef rstrErrDescr(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-03 22:31:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.10
'*************************************************************************
Public Function DeleteValue(ByVal mainKey As enumRegMainKey, _
                           ByVal subKey As String, _
                           ByVal keyV As String, _
                           Optional ByRef rlngErrNum As Long, _
                           Optional ByRef rstrErrDescr As String) As Boolean
   Dim hKey As Long
   
   On Error GoTo DeleteValueErr
   
   DeleteValue = False
   
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_WRITE, hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "ɾ��ע���ֵʱ����"
   End If
   
   If RegDeleteValue(hKey, keyV) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "ɾ��ע���ֵʱ����"
   End If
   
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "ɾ��ע���ֵʱ����"
   End If
   
   DeleteValue = True
   Err.Clear


DeleteValueErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function


'*************************************************************************
'**�� �� ����DeleteKey
'**��    �룺ByVal mainKey(enumRegMainKey)       -
'**        ��ByVal subKey(String)                -
'**        ��ByVal keyV(String)                  -
'**        ��Optional ByRef rlngErrNum(Long)     -
'**        ��Optional ByRef rstrErrDescr(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-03 22:31:47
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.10
'*************************************************************************
Public Function DeleteKey(ByVal mainKey As enumRegMainKey, _
                           ByVal subKey As String, _
                           ByVal keyV As String, _
                           Optional ByRef rlngErrNum As Long, _
                           Optional ByRef rstrErrDescr As String) As Boolean
   Dim hKey As Long
   On Error GoTo DeleteKeyErr
   
   DeleteKey = False
   
   If RegOpenKeyEx(mainKey, subKey, 0, KEY_WRITE, hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "ɾ��ע���ֵʱ����"
   End If
   
   If RegDeleteKey(hKey, keyV) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "ɾ��ע���ֵʱ����"
   End If
   
   If RegCloseKey(hKey) <> ERROR_SUCCESS Then
      Err.Raise vbObjectError + 1, , "ɾ��ע���ֵʱ����"
   End If
   
   DeleteKey = True
   Err.Clear

DeleteKeyErr:
   rlngErrNum = Err.Number
   rstrErrDescr = Err.Description
End Function

