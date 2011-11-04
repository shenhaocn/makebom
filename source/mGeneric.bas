Attribute VB_Name = "mGeneric"
'*************************************************************************************
'**ģ �� ����mGeneric
'**˵    ����TP-LINK SMB Switch Product Line Hardware Group ��Ȩ����2011 - 2012(C)
'**�� �� �ˣ�Shenhao
'**��    �ڣ�2011-10-31 23:36:58
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����ͨ��ģ���
'**��    ����V3.6.3
'*************************************************************************************
Option Explicit

'==============================================================================
' ��ȡ�ļ�����ʱ�� �޸�ʱ�� ����ʱ��
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
' Constant defining ( �������� )
'==============================================================================
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

'==============================================================================
' API function declare ( API�������� )
'==============================================================================
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = &H2 '������Ŀǰ�Ӵ�λ��
Const SWP_NOSIZE = &H1 '������Ŀǰ�Ӵ���С
Const HWND_TOPMOST = -1 '�趨Ϊ���ϲ�
Const HWND_NOTOPMOST = -2 'ȡ�����ϲ��趨
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE


'==============================================================================
Public ProcInfo As StatusBar
Private hFile As Long

'��ȡ�ļ�����ʱ��
Function GetFileCreateTime(filePath As String) As String

    Dim FileHandle As Long
    Dim fileinfo As BY_HANDLE_FILE_INFORMATION
    Dim lpReOpenBuff As OFSTRUCT, ft As SYSTEMTIME
    Dim tZone As TIME_ZONE_INFORMATION
    
    Dim dtCreate As Date ' ����ʱ�䡣

    Dim bias As Long
    ' ��ȡ���ļ���Handle��
    FileHandle = OpenFile(filePath, lpReOpenBuff, OF_READ)
    ' �����ļ���Handle��ȡ�ļ���Ϣ��
    Call GetFileInformationByHandle(FileHandle, fileinfo)
    Call CloseHandle(FileHandle)
    ' ��ȡʱ����Ϣ�� ��Ϊ��һ������ļ�ʱ���Ǹ�������ʱ�䡣
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias ' ʱ�� �Է�Ϊ��λ��
    Call FileTimeToSystemTime(fileinfo.ftCreationTime, ft) ' ת��ʱ��ṹ��
    dtCreate = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)
    
    GetFileCreateTime = CStr(dtCreate)

End Function

'��ȡ�ļ��޸�ʱ��
Function GetFileWriteTime(filePath As String) As String

    Dim FileHandle As Long
    Dim fileinfo As BY_HANDLE_FILE_INFORMATION
    Dim lpReOpenBuff As OFSTRUCT, ft As SYSTEMTIME
    Dim tZone As TIME_ZONE_INFORMATION
    
    Dim dtWrite As Date ' �޸�ʱ�䡣
    Dim bias As Long
    ' ��ȡ���ļ���Handle��
    FileHandle = OpenFile(filePath, lpReOpenBuff, OF_READ)
    ' �����ļ���Handle��ȡ�ļ���Ϣ��
    Call GetFileInformationByHandle(FileHandle, fileinfo)
    Call CloseHandle(FileHandle)
    ' ��ȡʱ����Ϣ�� ��Ϊ��һ������ļ�ʱ���Ǹ�������ʱ�䡣
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias ' ʱ�� �Է�Ϊ��λ��
    
    Call FileTimeToSystemTime(fileinfo.ftLastWriteTime, ft)
    dtWrite = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)

    GetFileWriteTime = CStr(dtWrite)

End Function

'��ȡ�ļ���ȡʱ��
Function GetFileAccessTime(filePath As String) As String

    Dim FileHandle As Long
    Dim fileinfo As BY_HANDLE_FILE_INFORMATION
    Dim lpReOpenBuff As OFSTRUCT, ft As SYSTEMTIME
    Dim tZone As TIME_ZONE_INFORMATION
    
    Dim dtAccess As Date ' ��ȡ���ڡ�
    Dim bias As Long
    ' ��ȡ���ļ���Handle��
    FileHandle = OpenFile(filePath, lpReOpenBuff, OF_READ)
    ' �����ļ���Handle��ȡ�ļ���Ϣ��
    Call GetFileInformationByHandle(FileHandle, fileinfo)
    Call CloseHandle(FileHandle)
    ' ��ȡʱ����Ϣ�� ��Ϊ��һ������ļ�ʱ���Ǹ�������ʱ�䡣
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias ' ʱ�� �Է�Ϊ��λ��
    
    Call FileTimeToSystemTime(fileinfo.ftLastAccessTime, ft)
    dtAccess = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)

    GetFileAccessTime = CStr(dtAccess)

End Function

'�� �����趨����Զ���������ϲ�
Function SetWindowsPos_TopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Function

'ȡ�����ϲ��趨
Function SetWindowsPos_NoTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Function

'��ȡ�ļ����� ����ȥ�����������
Function GetFileContents(filePath As String) As String
    Dim cfgfile            As String
    Dim cfgfilenum         As Integer
    Dim cfgfileContents    As String
    Dim cfgfileInfo()      As String
    Dim cleanContents      As String
    
    cfgfile = filePath
    
    cfgfilenum = FreeFile
    Open cfgfile For Binary As #cfgfilenum
        cfgfileContents = Space(LOF(cfgfilenum))
        Get #cfgfilenum, , cfgfileContents
    Close cfgfilenum

    cleanContents = UEFLoadTextFile(filePath, UEF_ANSI)
    
    'ȥ������
    'ɾ�����ļ��ж�ȡ�ĵ�����������Ŀ���
    Do While InStr(cleanContents, " " + vbCrLf) > 0
        '������з�ǰ�Ŀո�
        cleanContents = Replace(cleanContents, " " + vbCrLf, vbCrLf)
    Loop
    Do While InStr(1, cleanContents, vbCrLf + vbCrLf)
        '���������֮��Ŀ���
        cleanContents = Replace(cleanContents, vbCrLf + vbCrLf, vbCrLf)
    Loop
    If InStr(cleanContents, vbCrLf) = 1 Then
        '���Ϊ�׵Ŀ���
        cleanContents = Replace(cleanContents, vbCrLf, "", 1, 1)
    End If
    
    GetFileContents = cleanContents
End Function

Function GetBomContents(filePath As String) As String
    
    Dim BomLine()    As String
    Dim BomAtom()    As String
    Dim BomInfo      As String
    Dim AtomNum      As Integer
    Dim j            As Integer
    
    BomInfo = GetFileContents(filePath)
    
    '����BOM�ļ���˵��Ҫ�����¸�ʽ
    Do While InStr(1, BomInfo, vbCrLf + vbTab)
        BomInfo = Replace(BomInfo, vbCrLf + vbTab, vbTab)
    Loop
    
    '��ȡBOM�ļ�������
    BomLine = Split(BomInfo, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������
    AtomNum = UBound(Split(BomLine(0), vbTab)) '������������Ϊ��׼
    
    BomInfo = ""
    '�˶�ÿ�е�Ԫ�ظ��� Ԫ�ظ������� ���Բ���
     For j = 0 To UBound(BomLine) - 1
        BomAtom = Split(BomLine(j), vbTab)
        
        'Ԫ�ظ������� ����һ�кϲ�
        If UBound(BomAtom) <> AtomNum Then
            BomInfo = BomInfo + BomLine(j) + BomLine(j + 1) + vbCrLf
            j = j + 1
        Else
            BomInfo = BomInfo + BomLine(j) + vbCrLf
        End If
    Next j
    
    BomInfo = BomInfo + BomLine(j)
    
    'ȥ�����Ļ���
    'BomInfo = Left(BomInfo, Len(BomInfo) - 2)
    
    GetBomContents = BomInfo
    
'for debug
'    If Dir(SaveAsPath & "_test.txt") <> "" Then
'        Kill SaveAsPath & "_test.txt"
'    End If
'
'    Open SaveAsPath & "_test.txt" For Output As #1
'    Print #1, GetBomContents
'    Close #1
    
End Function

'����һ Ҫд����ļ���ַ�������� �޸ĵ����� �������� д����滻���ַ���
Public Function WriteTxt(strSourceFile As String, intRow As Long, StrLineNew As String)

    Dim StrOut As String, tmpStrLine As String
    Dim X As Long
    If Dir(strSourceFile) <> "" Then
        Open strSourceFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpStrLine
            X = X + 1
            If X = intRow Then tmpStrLine = StrLineNew
            If Not EOF(1) Then
                StrOut = StrOut & tmpStrLine & vbCrLf
            Else
                StrOut = StrOut & tmpStrLine
            End If
            'Debug.Print x
        Loop
        Close #1
    Else
        StrOut = StrLineNew
    End If
    
    Open strSourceFile For Output As #1
    Print #1, StrOut
    Close #1

End Function

'���� Ҫ������ı�������һ �ļ���ַ�������� ��ȡ������
Public Function ReadTxt(StrFile As String, intRow As Long) As String
    Dim StrOut As String, tmpStrLine As String
    Dim X As Long
    
    If Dir(StrFile, vbNormal) <> "" Then
        Open StrFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpStrLine
            X = X + 1
            If X = intRow Then ReadTxt = tmpStrLine: Exit Do
        Loop
        Close #1
    End If
    
End Function

'����Ϊ��ʾ���ݣ����ؽ��Ϊѡ���Ŀ¼
Function GetPath(Promt As String)
    Dim objshell
    Dim objfolder
    Set objshell = CreateObject("Shell.Application")
        Set objfolder = objshell.BrowseForFolder(0, Promt, 0, 0)
            If Not objfolder Is Nothing Then
                GetPath = objfolder.self.Path
            End If
        Set objfolder = Nothing
    Set objshell = Nothing
End Function

'����ִ�н��� ��������ģ�����ִ�н���
Function Process(ProcessNum As Integer, ProcessMsg As String)

    MainForm.StatusBar1.Panels(1) = ProcessMsg
    MainForm.StatusBar1.Panels(2) = ProcessNum & "%"
End Function

'ɾ���ɵ�Excel��ʽ�ļ�
Function KillExcel(ExcelFilePath As String)
On Error GoTo ErrorHandle
    If Dir(ExcelFilePath) <> "" Then
        Kill ExcelFilePath
    End If
    
    Exit Function
    
ErrorHandle:
    MsgBox "�ļ���" & vbCrLf & Right(ExcelFilePath, Len(ExcelFilePath) - InStrRev(ExcelFilePath, "\")) & vbCrLf & vbCrLf & _
           "�Ѿ��򿪻�ռ�ã��뽫��رպ��������г���", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "����"
    End
End Function


'*************************************************************************
'**�� �� ����AutoLoginERP
'**��    �룺uid(String) -
'**        ��pwd(String) -
'**��    ������
'**�����������Զ���¼ERPϵͳ Ϊ������ѯ��׼��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-01 00:44:04
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.7
'*************************************************************************
Function AutoLoginERP(uid As String, pwd As String)
    Dim i As Integer
    Dim j As Integer
    Dim orgName As String
    
    orgName = "���� �з�������ʦ"
    
    '��ȡ�������
    Dim SelStorage As String
    
    SelStorage = GetRegValue(App.EXEName, "Storage", "TP1")
    
    Select Case SelStorage
    Case "TP1"
        orgName = "���� �з�������ʦ"
    Case "TP2"
        orgName = "���� �з�������ʦ"
    Case "TP3"
        orgName = "OEM �з�������ʦ"
    Case Else
        
        orgName = "���� �з�������ʦ"
    End Select
        
    'MsgBox "���ڴ�ERPϵͳ�����Ե�..."
    
    With CreateObject("InternetExplorer.Application")
        .Visible = False
        .Navigate "http://erpprod.tplink.net:8007/OA_HTML/AppsLocalLogin.jsp?cancelUrl=/OA_HTML/AppsLocalLogin.jsp&langCode=ZHS"
        Do Until .ReadyState = 4
            DoEvents
        Loop
        .Document.Forms(0).All("username").Value = uid
        .Document.Forms(0).All("password").Value = pwd
        .Document.Forms(0).submit
        '�Զ���¼����
        
        Do While .busy Or .ReadyState <> 4
        Loop

        For i = 0 To .Document.All.Length - 1
            If (LCase(.Document.All(i).tagname)) = "a" Then
                If InStr(.Document.All(i).innerText, orgName) <> 0 Then
                    .Document.All(i).Click
                    
                    Do While .busy Or .ReadyState <> 4
                    Loop
                    
                    '���뵽������ѯ
                    For j = 0 To .Document.All.Length - 1
                        If (LCase(.Document.All(j).tagname)) = "a" Then
                            If InStr(.Document.All(j).innerText, "����������Դ��ѯ") <> 0 Then
                                .Document.All(j).Click
                                Exit Function
                            End If
                        End If
                    Next j
                    
                End If
            End If
        Next i
        
    End With
    
End Function


'*************************************************************************
'**�� �� ����GetInfoFromERP
'**��    �룺uid(String) -
'**        ��pwd(String) -
'**��    ������
'**������������ERPϵͳ�л�ȡ��Ϣ����Functionδ���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-01 00:44:33
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.7
'*************************************************************************
Function GetInfoFromERP(uid As String, pwd As String)
    Dim i As Integer
    Dim j As Integer
    Dim orgName As String
    
    orgName = "���� �з�������ʦ"
    
    '��ȡ�������
    Dim SelStorage As String
    
    SelStorage = GetRegValue(App.EXEName, "Storage", "TP1")
    
    Select Case SelStorage
    Case "TP1"
        orgName = "���� �з�������ʦ"
    Case "TP2"
        orgName = "���� �з�������ʦ"
    Case "TP3"
        orgName = "OEM �з�������ʦ"
    Case Else
        
        orgName = "���� �з�������ʦ"
    End Select
        
    'MsgBox "���ڴ�ERPϵͳ�����Ե�..."
    
    With CreateObject("InternetExplorer.Application")
        .Visible = True
        .Navigate "http://erpprod.tplink.net:8007/OA_HTML/AppsLocalLogin.jsp?cancelUrl=/OA_HTML/AppsLocalLogin.jsp&langCode=ZHS"
        Do Until .ReadyState = 4
            DoEvents
        Loop
        .Document.Forms(0).All("username").Value = uid
        .Document.Forms(0).All("password").Value = pwd
        .Document.Forms(0).submit
        '�Զ���¼����
        
        Do While .busy Or .ReadyState <> 4
        Loop

        For i = 0 To .Document.All.Length - 1
            If (LCase(.Document.All(i).tagname)) = "a" Then
                If InStr(.Document.All(i).innerText, orgName) <> 0 Then
                    .Document.All(i).Click
                    
                    Do While .busy Or .ReadyState <> 4
                    Loop
                    
                    '���뵽������ѯ
                    For j = 0 To .Document.All.Length - 1
                        If (LCase(.Document.All(j).tagname)) = "a" Then
                            If InStr(.Document.All(j).innerText, "����ϲ�ѯ") <> 0 Then
                                .Document.All(j).Click
                                GoTo Check_Done
                            End If
                        End If
                    Next j
                    
                End If
            End If
        Next i
        
Check_Done:
        '���뵽����ϲ�ѯλ�� �ϴ�Excel ������ѯ�����
        Do While .busy Or .ReadyState <> 4
        Loop
        
        '<option value="222">TP4</option>
        '<option value="262">TP5</option>
        '<option value="382">TP7</option>
        '<option value="121">TP1</option>
        '<option value="122">TP2</option>
        '<option value="123">TP3</option>
        '<option value="442">TP8</option>
        '<option value="321">TP6</option>
        .Document.Forms(0).All("OrganizationCode").Value = "121"
        '<option value="L">��������</option>
        '<option value="S">��������</option>
        .Document.GetElementById("QueryType").Value = "L"
        .Document.GetElementById("QueryType").FireEvent "onchange"
        .Document.GetElementById("QueryType").FireEvent "onclick"
        .Document.GetElementById("QueryType").FireEvent "onblur"
        
        '���뵽����ϲ�ѯλ�� �ϴ�Excel ������ѯ�����
        Do While .busy Or .ReadyState <> 4
        Loop
        
        'VB.Clipboard.SetText "E:\��������\MakeBomTest\test\TL-SL1210-2.0\BOM\TL-SL1210(UN)_2.0_SCHV_2.0_DEV1_PCBV_1.0SP1_������Դ��ѯ.xls"
        
        .Document.GetElementById("UploadFile_oafileUpload").Focus
        .Document.GetElementById("UploadFile_oafileUpload").Click
        DoEvents
        
        'UploadFile
        
        '.Document.Forms(0).All("UploadFile_oafileUpload").Value = "E:\��������\MakeBomTest\test\TL-SL1210-2.0\BOM\TL-SL1210(UN)_2.0_SCHV_2.0_DEV1_PCBV_1.0SP1_������Դ��ѯ.xls"
        
        .Document.GetElementById("Go").FireEvent "onclick"
        
        '���뵽����ϲ�ѯλ�� �ϴ�Excel ������ѯ�����
        Do While .busy Or .ReadyState <> 4
        Loop
        
        '���¿�ʼ��ѯ
        
        
        
    End With
    
End Function

Function UploadFile()
    Dim ParentWnd As Long   '�����ھ��
    Dim ClientWnd As Long   '�Ӵ��ھ��
    
    Dim msgstr    As String
    
    VB.Clipboard.SetText "E:\��������\MakeBomTest\test\TL-SL1210-2.0\BOM\TL-SL1210(UN)_2.0_SCHV_2.0_DEV1_PCBV_1.0SP1_������Դ��ѯ.xls"
    
    '������ȡ����ָ�����ھ�����̣�ע���޸������ʹ�����
    ParentWnd = FindWindow("Dialog", "ѡ��Ҫ���ص��ļ�")
    If ParentWnd = 0 Then
        'MsgBox "û���ҵ�������", 16, "����"
        Exit Function
    Else
        MsgBox "�ҵ�����"
        SendKeys "^v" & "~"
    End If
    
    
    'SendKeys "^v"
    
    'ȡ�õ�һ���Ӵ��ڵľ��
    'ClientWnd = GetWindow(ParentWnd, GW_CHILD)
    'If ClientWnd = 0 Then
    '    MsgBox "��ָ��������û�з����Ӵ��ڵĴ���", 16, "����"
    '    Exit Function
    'End If
    
    '��ʼѭ������������ͬ��ε��Ӵ���
    'Do
    '    DoEvents
    '    msgstr = msgstr & "�Ӵ��ڣ�" & ClientWnd & vbCrLf
    '    ClientWnd = GetWindow(ClientWnd, GW_HWNDNEXT)
    'Loop While ClientWnd <> 0
    
    'MsgBox "��ɴ���", 64, "��ʾ"
    
End Function

