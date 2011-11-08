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


Public ProcInfo As StatusBar

'*************************************************************************
'**�� �� ����Process
'**��    �룺ProcessNum(Integer) -
'**        ��ProcessMsg(String)  -
'**��    ������
'**��������������ִ�н��� ��������ģ�����ִ�н���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:00:01
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
Function Process(ProcessNum As Integer, ProcessMsg As String)

    MainForm.StatusBar1.Panels(1) = ProcessMsg
    MainForm.StatusBar1.Panels(2) = ProcessNum & "%"
End Function

'*************************************************************************
'**�� �� ����SetWindowsPos_TopMost
'**��    �룺hwnd(Long) -
'**��    ������
'**������������ �����趨����Զ���������ϲ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:00:17
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
Function SetWindowsPos_TopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Function

'*************************************************************************
'**�� �� ����SetWindowsPos_NoTopMost
'**��    �룺hwnd(Long) -
'**��    ������
'**����������ȡ�����ϲ��趨
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:00:33
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
Function SetWindowsPos_NoTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Function


'*************************************************************************
'**�� �� ����SetRegValue
'**��    �룺AppName(String)            -
'**        ��KeyName(String)            -
'**        ��ByVal lType(enumRegSzType) -
'**        ��                           -
'**��    ����(Boolean) -
'**��������������ע�������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:00:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
Public Function SetRegValue(AppName As String, KeyName As String, ByVal lType As enumRegSzType, ByVal KeyValue) As Boolean

     SetRegValue = SetValue(iHKEY_CURRENT_USER, "SOFTWARE\" + AppName, KeyName, lType, KeyValue)
     
End Function


'*************************************************************************
'**�� �� ����SetRegValue
'**��    �룺AppName(String)            -
'**        ��KeyName(String)            -
'**        ��ByVal lType(enumRegSzType) -
'**        ��                           -
'**��    ����(Boolean) -
'**������������ȡע�������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:00:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
Public Function GetRegValue(AppName As String, KeyName As String, Optional DefaultKeyValue As Variant) As Variant

    If GetValue(iHKEY_CURRENT_USER, "SOFTWARE\" + AppName, KeyName, GetRegValue) = False Then
        GetRegValue = DefaultKeyValue
    End If
    
End Function

Public Function GetExcelVer() As String

    Dim sKeyReg() As String
    Dim tmpReg()  As String
    Dim Cnt       As Long
    
    
    If RegEnumKeyVal(iHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office", sKeyReg) = False Then
        GetExcelVer = ""
        Exit Function
    End If

    For Cnt = 0 To UBound(sKeyReg)
        Select Case sKeyReg(Cnt)
        Case "14.0", "12.0", "11.0", "10.0", "9.0", "8.0", "5.0", "4.0", "3.0" 'Excel�����汾��
            If RegEnumKeyVal(iHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\" & sKeyReg(Cnt) & "\Excel", tmpReg) = True Then
                GetExcelVer = sKeyReg(Cnt)
                Exit For
            End If

        Case Else
            If CInt(sKeyReg(Cnt)) > 14 Then
                GetExcelVer = sKeyReg(Cnt)
            Else
                GetExcelVer = ""
            End If
        End Select
    Next
    
End Function

'ȥ���������ж����Chr(0)
Public Function StripTerminator(sInput As String) As String
  Dim ZeroPos As Integer
  '�ҵ���һ��Chr(0)
  ZeroPos = InStr(1, sInput, vbNullChar)
  If ZeroPos > 0 Then '�������,��ȥ���������е�����
    StripTerminator = Left$(sInput, ZeroPos - 1)
  Else '���������,�����κβ���
    StripTerminator = sInput
  End If
End Function



'*************************************************************************
'**�� �� ����KillExcel
'**��    �룺ExcelFilePath(String) -
'**��    ������
'**����������ɾ���ɵ�Excel��ʽ�ļ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:01:29
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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
'**�� �� ����GetPath
'**��    �룺Promt(String) - ��ʾ����
'**��    �������Ϊѡ���Ŀ¼
'**����������ѡ��Ŀ¼·��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:01:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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

'*************************************************************************
'**�� �� ����GetFileContents
'**��    �룺filePath(String) -
'**��    ����(String) -
'**������������ȡ�ļ����� ����ȥ�����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:02:21
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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


'*************************************************************************
'**�� �� ����GetBomContents
'**��    �룺filePath(String) -
'**��    ����(String) -
'**������������ȡ����BOM�ļ�����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:02:30
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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


'*************************************************************************
'**�� �� ����WriteTxt
'**��    �룺strSourceFile(String) - Ҫд����ļ���ַ
'**        ��intRow(Long)          - �޸ĵ�����
'**        ��StrLineNew(String)    - д����滻���ַ���
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:02:52
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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

'*************************************************************************
'**�� �� ����ReadTxt
'**��    �룺StrFile(String) -  �ļ���ַ
'**        ��intRow(Long)    -  ��ȡ������
'**��    ����(String) -  Ҫ������ı�
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:03:13
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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
        
        '.Document.Forms(0).All("UploadFile_oafileUpload").Value = "E:\��������\MakeBomTest\test\TL-SL1210-2.0\BOM\TL-SL1210(UN)_2.0_SCHV_2.0_DEV1_PCBV_1.0SP1_������Դ��ѯ.xls"
        
        .Document.GetElementById("Go").FireEvent "onclick"
        
        '���뵽����ϲ�ѯλ�� �ϴ�Excel ������ѯ�����
        Do While .busy Or .ReadyState <> 4
        Loop
        
        '���¿�ʼ��ѯ
        
    End With
    
End Function


'*************************************************************************
'**�� �� ����GetFileCreateTime
'**��    �룺filePath(String) -
'**��    ����(String) -
'**������������ȡ�ļ�����ʱ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:03:38
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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


'*************************************************************************
'**�� �� ����GetFileWriteTime
'**��    �룺filePath(String) -
'**��    ����(String) -
'**������������ȡ�ļ��޸�ʱ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:03:46
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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


'*************************************************************************
'**�� �� ����GetFileAccessTime
'**��    �룺filePath(String) -
'**��    ����(String) -
'**������������ȡ�ļ���ȡʱ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-11-07 22:03:55
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.16
'*************************************************************************
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

