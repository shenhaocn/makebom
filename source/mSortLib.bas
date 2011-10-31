Attribute VB_Name = "mSortLib"
'*************************************************************************************
'**ģ �� ����mSortLib
'**˵    ����TP-LINK SMB Switch Product Line Hardware Group ��Ȩ����2011 - 2012(C)
'**�� �� �ˣ�Shenhao
'**��    �ڣ�2011-10-31 23:38:03
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����
'**��    ����V3.6.3
'*************************************************************************************
Option Explicit

'�����
Public Enum MountType
LIB_LEAD = 0 'LEAD Type
LIB_SMD      'smd type
LIB_NONE     'None type
End Enum

Public LibFilePath     As String  '���ļ�·����Ϣ


'*************************************************************************
'**�� �� ����InitLib
'**��    �룺filePath(String) -
'**��    ����(Boolean) -
'**������������ʼ����װ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-31 23:41:31
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.3
'*************************************************************************
Function InitLib(filePath As String) As Boolean
    '�Ƿ���ڿ��ļ�
    If Dir(filePath) = "" Then
         MsgBox "���ļ������ڣ�", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "���ش���"
         InitLib = False
         Exit Function
    End If
    
    LibFilePath = filePath
    
    '������ļ��������Ժ���ȷ��
    Dim LibLine()          As String
    Dim LibAtom()          As String
    Dim i                  As Integer
    
    LibLine = OpenLibs()
    For i = 2 To UBound(LibLine) - 1
        LibAtom = Split(LibLine(i), Space(1))
        If UBound(LibAtom) <> 3 Then
            MsgBox "���ļ�STD.lst����������������STD.lst���ļ���", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            End
        End If
    Next i
    
    InitLib = True
End Function


'*************************************************************************
'**�� �� ����OpenLibs
'**��    �룺��
'**��    �����ַ�����������һ���ո�ָ�������Ե��ַ��� �Ը�������Ϊ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-31 23:42:02
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.3
'*************************************************************************
Function OpenLibs() As String()
    Dim FileContents    As String
    Dim fileinfo()      As String
        
    FileContents = GetFileContents(LibFilePath)
    fileinfo = Split(FileContents, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������

    Dim j As Integer
    For j = 2 To UBound(fileinfo) - 1
        Do While InStr(fileinfo(j), Space(2))
            fileinfo(j) = Replace(fileinfo(j), Space(2), Space(1)) '�������Ŀո�
        Loop
    Next j
    
    '���ص��ַ�����������һ���ո�ָ�������Ե��ַ��� �Ը�������Ϊ����
    OpenLibs = fileinfo
        
End Function


'*************************************************************************
'**�� �� ����GetLibsVersion
'**��    �룺��
'**��    ����(String) -
'**������������ȡ��װ��汾��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-31 23:42:48
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.3
'*************************************************************************
Function GetLibsVersion() As String
    Dim FileContents    As String
    Dim fileinfo()      As String
        
    FileContents = GetFileContents(LibFilePath)
    fileinfo = Split(FileContents, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������
    
    '�汾��Ϣ�ڵ�1��
    GetLibsVersion = Right(fileinfo(0), Len(fileinfo(0)) - InStrRev(fileinfo(0), "VERSION:") + 1)
    
    GetLibsVersion = Replace(GetLibsVersion, "VERSION:" & Space(2), "")
        
End Function


'*************************************************************************
'**�� �� ����ReadLibs
'**��    �룺Lib(MountType) As String() -
'**��    ������
'**������������ȡָ���ķ�װ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-31 23:43:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.3
'*************************************************************************
Function ReadLibs(Lib As MountType) As String()
    Dim LibsInfo()      As String
    Dim Libstr()        As String
    LibsInfo = OpenLibs()
    
    'ͳ�ƿ��и��з�װ���͵�����
    Dim j               As Integer
    Dim leadNum         As Integer
    Dim smdNum          As Integer
    Dim otherNum        As Integer
    
    leadNum = 0
    smdNum = 0
    otherNum = 0
    For j = 2 To UBound(LibsInfo) - 1
        Libstr = Split(LibsInfo(j), " ")
        If UBound(Libstr) = 3 Then
            Select Case Libstr(3)
            Case "L"
                leadNum = leadNum + 1
            Case "S"
                smdNum = smdNum + 1
            Case "S+"
                leadNum = leadNum + 1
                smdNum = smdNum + 1
            Case "N"
                otherNum = otherNum + 1
            Case Else
                MsgBox "��װ [" & Libstr(0) & "]" & "δ֪����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "���ļ�����"
            End Select
        End If
    Next j
    
    Dim LibNum As Integer
    Dim cfgInfo() As String
    
    Select Case Lib
    Case LIB_LEAD
        LibNum = leadNum
        ReDim cfgInfo(leadNum) As String
    Case LIB_SMD
        LibNum = smdNum
        ReDim cfgInfo(smdNum) As String
    Case LIB_NONE
        LibNum = otherNum
        ReDim cfgInfo(otherNum) As String
    Case Else
        ' SHOULD NEVER BE HERE
    End Select
    
    Dim i As Integer
    
    i = 0
    For j = 2 To UBound(LibsInfo) - 1
        Libstr = Split(LibsInfo(j), " ")
        If UBound(Libstr) = 3 Then
            Select Case Libstr(3)
            Case "L"
                If Lib = LIB_LEAD Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case "S"
                If Lib = LIB_SMD Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case "S+"
                If Lib = LIB_LEAD Or Lib = LIB_SMD Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case "N"
                If Lib = LIB_NONE Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case Else
                MsgBox "��װ [" & Libstr(0) & "]" & "δ֪����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "���ļ�����"
            End Select
        End If
    Next j

    ReadLibs = cfgInfo
        
End Function


'*************************************************************************
'**�� �� ����QueryLib
'**��    �룺LibInfo()(String) -
'**        ��QueryStr(String)  -
'**��    ����(Integer) -
'**������������ѯ��װ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-31 23:43:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.3
'*************************************************************************
Function QueryLib(LibInfo() As String, QueryStr As String) As Integer
    Dim i As Integer
    QueryLib = 0
    For i = 0 To UBound(LibInfo) - 1
        If QueryStr <> "" And LibInfo(i) <> "" Then
            If QueryStr = LibInfo(i) Then
                QueryLib = 1
                Exit For
            End If
        End If
    Next i

End Function

