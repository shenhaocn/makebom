Attribute VB_Name = "mInitBOM"
Option Explicit

Public BomItemNumber   As Integer 'BomԪ�ض�λ��Ϣ
Public BomPartNumber   As Integer
Public BomValue        As Integer
Public BomQuantity     As Integer
Public BomPartRef      As Integer
Public BomPCBfootprint As Integer

Public PartNum(6)      As Integer '����Ԫ������Ϣ

Public ProjectDir      As String  '�����ϴδ򿪵�Ŀ¼
Public ItemName        As String  '�����ϴδ򿪵�Ŀ¼

Public BomFilePath     As String  'ԭʼ�ļ�������
Public SaveAsPath      As String  'BOM������ļ�·��
Public tsvFilePath     As String  'tsv�ļ�·����Ϣ

Function BuildProjectPath(srcPath As String)
    '��������������Ҫ��Ŀ¼��Ϣ�������������У����˿�д����Щ·��
    Dim tmpPath As String
    BomFilePath = srcPath
    tmpPath = Right(BomFilePath, Len(BomFilePath) - InStrRev(BomFilePath, "\"))
    ProjectDir = Left(BomFilePath, InStrRev(BomFilePath, "\") - 1) & "\"
    tmpPath = ProjectDir & "BOM\" & tmpPath
    SaveAsPath = Left(tmpPath, InStrRev(tmpPath, ".") - 1)
    '�ڹ���Ŀ¼�´���BOMĿ¼
    If Dir(ProjectDir & "BOM\") = "" Then
        MkDir ProjectDir & "BOM\"
    End If
    SaveSetting App.EXEName, "ProjectDir", "�ϴι���Ŀ¼", ProjectDir
End Function

Function ClearPath()
    '��������������Ҫ��Ŀ¼��Ϣ�������������У����˿�д����Щ·��
    BomFilePath = ""
    SaveAsPath = ""
    tsvFilePath = ""
End Function

Function ReadBomFile() As Boolean
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim newbomstr()     As String
    
    Dim filenum         As Integer
    Dim i               As Integer
    
    filenum = FreeFile
    Open BomFilePath For Binary As #filenum
        FileContents = Space(LOF(filenum))
        Get #filenum, , FileContents
    Close filenum
    fileinfo = Split(FileContents, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������
    
    'j��ʾԴ�ļ�BOM�е���
    'i��ʾ�е�ĳһ�У���tab�ָ�ģ�
    'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint
    '0-----------1-----------2-------3-----------4---------------5------------
    'ע��orCAD���������п��ܲ�������һ��  �����Ҫ��λ
    Process 3, "��ȡ.BOM�ļ���Ϣ..."
    
    BomItemNumber = -1
    BomPartNumber = -1
    BomValue = -1
    BomQuantity = -1
    BomPartRef = -1
    BomPCBfootprint = -1
    
    newbomstr = Split(fileinfo(0), vbTab)
    For i = 0 To UBound(newbomstr)
        If newbomstr(i) = "Item Number" Then
            BomItemNumber = i
        End If
        If newbomstr(i) = "Part Number" Then
            BomPartNumber = i
        End If
        If newbomstr(i) = "Value" Then
            BomValue = i
        End If
        If newbomstr(i) = "Quantity" Then
            BomQuantity = i
        End If
        If newbomstr(i) = "Part Reference" Then
            BomPartRef = i
        End If
        If newbomstr(i) = "PCB Footprint" Then
            BomPCBfootprint = i
        End If
        'If BomItemNumber > 5 Or BomPartNumber > 5 Or BomValue > 5 Or BomQuantity > 5 Or BomPartRef > 5 Or BomPCBfootprint > 5 Then
        '    MsgBox "BOM�ļ���ʽ����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
        '    ReadBomFile = False
        '    Exit Function
        'End If
    Next
    
    If BomItemNumber = -1 Or BomPartNumber = -1 Or BomValue = -1 Or BomQuantity = -1 Or BomPartRef = -1 Or BomPCBfootprint = -1 Then
        MsgBox "BOM�ļ���ʽ����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
        ReadBomFile = False
        Exit Function
    End If
    
    Dim IgLibInfo()        As String
    IgLibInfo = ReadLibs(LIB_NONE)
    
    Dim IsNone As Integer
    Dim j      As Integer
    For j = 1 To UBound(fileinfo) - 1
        newbomstr = Split(fileinfo(j), vbTab)
        'BOM��ÿ����"N"��Ԫ������ӵ���Ϻ�(��Ϊģ���Ϻ�)
        IsNone = QueryLib(IgLibInfo, newbomstr(BomPCBfootprint))
        If IsNone = 0 Then
            If newbomstr(BomPartNumber) = "" Then
                ReadBomFile = False
                MsgBox "��װΪ[" & newbomstr(BomPCBfootprint) & "]�ϺŲ����ڣ�", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "BOM�ļ��淶����"
                Exit Function
            End If
        End If
    Next
    
    ReadBomFile = True
End Function


Function CalcPartNum()

    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook, PLxlBook As Excel.Workbook, NBxlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet, PLxlSheet As Excel.Worksheet, NBxlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set xlSheet = xlBook.Worksheets(1)
    
    Set NBxlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Set NBxlSheet = NBxlBook.Worksheets(1)

    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    
    Dim rngNC           As Range
    Dim rngDB           As Range
    Dim rngDBNC         As Range
    
    Dim rngEND          As Range
    
    '��λ����Ԫ��λ��
    With xlSheet.Cells
        Set rngSMT = .Find("SMTԪ��", lookin:=xlValues)
        Set rngLEAD = .Find("DIPԪ��", lookin:=xlValues)
        Set rngOther = .Find("����Ԫ��", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOMģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    '����Ԫ������
    SmtPartNum = rngLEAD.Row - rngSMT.Row - 1
    LeadPartNum = rngOther.Row - rngLEAD.Row - 1
    OtherPartNum = rngEND.Row - rngOther.Row - 1
    
    If SmtPartNum = 1 And xlSheet.Cells(rngSMT.Row + 1, 5) = "" Then
        SmtPartNum = 0
    End If
    If LeadPartNum = 1 And xlSheet.Cells(rngLEAD.Row + 1, 5) = "" Then
        LeadPartNum = 0
    End If
    If OtherPartNum = 1 And xlSheet.Cells(rngOther.Row + 1, 5) = "" Then
        OtherPartNum = 0
    End If

    '��λ����Ԫ��λ��
    With NBxlSheet.Cells
        Set rngNC = .Find("NCԪ��", lookin:=xlValues)
        Set rngDB = .Find("DBGԪ��", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NCԪ��", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Or rngEND Is Nothing Then
            MsgBox "NC_DBGģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    Dim NcPartNum       As Integer
    Dim DbgPartNum      As Integer
    Dim DbNcPartNum     As Integer
    
    NcPartNum = rngDB.Row - rngNC.Row - 1
    DbgPartNum = rngDBNC.Row - rngDB.Row - 1
    DbNcPartNum = rngEND.Row - rngDBNC.Row - 1
    

    If NcPartNum = 1 And NBxlSheet.Cells(rngNC.Row + 1, 5) = "" Then
        NcPartNum = 0
    End If
    If DbgPartNum = 1 And NBxlSheet.Cells(rngDB.Row + 1, 5) = "" Then
        DbgPartNum = 0
    End If
    If DbNcPartNum = 1 And NBxlSheet.Cells(rngDBNC.Row + 1, 5) = "" Then
        DbNcPartNum = 0
    End If

    
    '����Ԫ������������������
    PartNum(0) = NcPartNum
    PartNum(1) = DbgPartNum
    PartNum(2) = DbNcPartNum
    
    PartNum(3) = LeadPartNum
    PartNum(4) = SmtPartNum
    PartNum(5) = OtherPartNum
    
    Dim msgstr As String
    msgstr = "Ԫ����Ϣ��ȡ�ɹ���" & vbCrLf & vbCrLf
    msgstr = msgstr + "          ��װ   Ԫ������Ϊ   �� " & SmtPartNum & vbCrLf
    msgstr = msgstr + "          ��װ   Ԫ������Ϊ   �� " & LeadPartNum & vbCrLf
    msgstr = msgstr + "          ����   Ԫ������Ϊ   �� " & OtherPartNum & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     Ԫ������Ϊ   �� " & NcPartNum & vbCrLf
    msgstr = msgstr + "          DBG    Ԫ������Ϊ   �� " & DbgPartNum & vbCrLf
    msgstr = msgstr + "          DBG_NC Ԫ������Ϊ   �� " & DbNcPartNum & vbCrLf & vbCrLf
    msgstr = msgstr + "          ��ѡ��.tsv�ļ�·�����������" & vbCrLf & vbCrLf
    msgstr = msgstr + "    ע�⣺���ɵ�PCBA_BOM�ļ���Ҫ����޸ĺ�ſɹ����� "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "Ԫ����Ϣ"
    
    xlBook.Close (True) '�رչ�����
    NBxlBook.Close (True)
    
ErrorHandle:
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
End Function

Function BomDraft()
    On Error GoTo ErrorHandle
    
    Process 4, "����Excel��ʽ BOM�ĵ�..."
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook, PLxlBook As Excel.Workbook, NBxlBook As Excel.Workbook, NonexlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet, PLxlSheet As Excel.Worksheet, NBxlSheet As Excel.Worksheet, NonexlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Process 5, "��PCBA_BOM����..."
        
    'PCBA_BOM
    Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set xlSheet = xlBook.Worksheets(1)
    
    Process 6, "��������ѯ�ļ�..."
    '��������Դ��ѯxls
    Set PLxlBook = xlApp.Workbooks.Open(SaveAsPath & "_������Դ��ѯ.xls")
    Set PLxlSheet = PLxlBook.Worksheets(1)
    
    Process 7, "��NC_DBGԪ���ĵ�..."
    '��NC_DBGԪ��xls
    Set NBxlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Set NBxlSheet = NBxlBook.Worksheets(1)
    
    '��NoneԪ����
    Set NonexlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
    Set NonexlSheet = NonexlBook.Worksheets(1)
    
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    
    Dim rngNC           As Range
    Dim rngDB           As Range
    Dim rngDBNC         As Range
    
    Process 8, "��λPCBA_BOM�и�Ԫ����ʼλ����Ϣ..."
    '��λ����Ԫ��λ��
    With xlSheet.Cells
        Set rngSMT = .Find("SMTԪ��", lookin:=xlValues)
        Set rngLEAD = .Find("DIPԪ��", lookin:=xlValues)
        Set rngOther = .Find("����Ԫ��", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Then
            MsgBox "PCBA_BOMģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    Process 9, "��λNC_DBG�и�Ԫ����ʼλ����Ϣ..."
    '��λ����Ԫ��λ��
    With NBxlSheet.Cells
        Set rngNC = .Find("NCԪ��", lookin:=xlValues)
        Set rngDB = .Find("DBGԪ��", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NCԪ��", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Then
            MsgBox "NC_DBGģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    '========================================================
    '��ȡ����Ϣ
    'LEAD ��
    Process 10, "��ȡ���ļ���Ϣ..."
    
    Dim leadLibInfo()      As String
    Dim smtLibInfo()       As String
    Dim IgLibInfo()        As String
    
    leadLibInfo = ReadLibs(LIB_LEAD)
    smtLibInfo = ReadLibs(LIB_SMD)
    IgLibInfo = ReadLibs(LIB_NONE)
        
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim bomstr()        As String
    Dim strtmp          As String
    
    Dim NcPartNum       As Integer
    Dim DbgPartNum      As Integer
    Dim DbNcPartNum     As Integer
    Dim NonePartNum     As Integer
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    
    Dim PLPartNum       As Integer
    
    Dim IsLead          As Integer
    Dim IsSmt           As Integer
    Dim IsNone          As Integer
    
    FileContents = GetFileContents(BomFilePath)
    fileinfo = Split(FileContents, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������
    
    Dim j               As Integer
    Dim k               As Integer
    
    On Error GoTo ErrorHandle
    For j = 1 To UBound(fileinfo) - 1
        
        bomstr = Split(fileinfo(j), vbTab)
            
         '�Ϻź�������������ϻ���
        If UBound(bomstr) < 5 Then
            strtmp = fileinfo(j)
            For k = j + 1 To UBound(fileinfo)
                If Len(fileinfo(k)) > 1 Then
                    bomstr = Split(strtmp & fileinfo(k), vbTab)
                    j = k
                    Exit For
                End If
            Next
        End If
        
        Process j * 40 / UBound(fileinfo) + 10, "������װ[" & bomstr(BomPCBfootprint) & "]..."
        
        '�����Ϻ� Ҫ���������������ѯ��Excel��
        
        If IsNumeric(bomstr(BomPartNumber)) = True And bomstr(BomPartNumber) <> "" Then
            PLxlSheet.Cells(PLPartNum + 1, 1) = bomstr(BomPartNumber)
            PLPartNum = PLPartNum + 1
        End If
        
        If InStr(bomstr(BomValue), "_DBG_NC") > 0 Or bomstr(BomValue) = "DBG_NC" Then
            'DBG_NCԪ��
            DbNcPartNum = DbNcPartNum + 1
            xlsInsert NBxlSheet, DbNcPartNum, rngDBNC.Row, bomstr
            
        ElseIf InStr(bomstr(BomValue), "_DBG") > 0 Or bomstr(BomValue) = "DBG" Then
            'DBGԪ��
            DbgPartNum = DbgPartNum + 1
            xlsInsert NBxlSheet, DbgPartNum, rngDB.Row, bomstr
           
        ElseIf InStr(bomstr(BomValue), "_NC") > 0 Or bomstr(BomValue) = "NC" Then
            'NCԪ��
            NcPartNum = NcPartNum + 1
            xlsInsert NBxlSheet, NcPartNum, rngNC.Row, bomstr
            
        Else
        
            If bomstr(BomPCBfootprint) = "" Then
                MsgBox bomstr(BomPartNumber) & "��PCB footprintΪ��", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "����"
                GoTo ErrorHandle
            End If
            
            '========================================================
            '��ͨԪ�� ����Ԫ����װ����
            '�ж�Ԫ������
            IsLead = QueryLib(leadLibInfo, bomstr(BomPCBfootprint))
            IsSmt = QueryLib(smtLibInfo, bomstr(BomPCBfootprint))
            IsNone = QueryLib(IgLibInfo, bomstr(BomPCBfootprint))
                
            If IsLead = 1 And IsSmt = 0 And IsNone = 0 Then
                'ͳ�Ʋ�д��LEADԪ��
                LeadPartNum = LeadPartNum + 1
                xlsInsert xlSheet, LeadPartNum, rngLEAD.Row, bomstr
                
            ElseIf IsLead = 0 And IsSmt = 1 And IsNone = 0 Then
                'ͳ�Ʋ�д��SMTԪ��
                SmtPartNum = SmtPartNum + 1
                xlsInsert xlSheet, SmtPartNum, rngSMT.Row, bomstr
            
            ElseIf IsLead = 0 And IsSmt = 0 And IsNone = 1 Then
                'ͳ�Ʋ�д�뵥�����ļ��� NoneԪ��
                NonePartNum = NonePartNum + 1
                '����ʹ�õ���NC_DBGģ�棬��˿���rngNC.Row,������Ҫ���¶�λ
                xlsInsert NonexlSheet, NonePartNum, rngNC.Row, bomstr
                
            ElseIf IsLead = 1 And IsSmt = 1 And IsNone = 0 Then
                '����SMTԪ�� �������� ������ɫ��ʾ
                SmtPartNum = SmtPartNum + 1
                xlsInsert xlSheet, SmtPartNum, rngSMT.Row, bomstr
                xlSheet.Rows(rngSMT.Row & ":" & rngSMT.Row).Interior.Color = 16737792
                
            Else
               '���ļ���û�в鵽��װ���ܾ�����BOM
                MsgBox "��װ[" & bomstr(BomPCBfootprint) & "]�������ڿ��ļ��У�����¿��ļ���"
                OtherPartNum = OtherPartNum + 1
                xlsInsert xlSheet, OtherPartNum, rngOther.Row, bomstr
                GoTo ErrorHandle
            End If
            
            IsLead = 0
            IsSmt = 0
            IsNone = 0

        End If
    Next j
    
    '�޸Ļ�������
    xlSheet.Cells(2, 1) = "���ͣ�  " & MainForm.ItemNameText.Text & "            PCBA �汾��                       ���Ʒ��ţ�"
    If MainForm.ItemNameText.Text = "" Then
        xlSheet.Cells(2, 1).Font.ColorIndex = 5
    End If
     
    Process 50, "������ѯ�ļ��������..."
    
    Dim msgstr As String
    msgstr = ""
    msgstr = msgstr + "          ��������ѯ��Ԫ�������� " & PLPartNum & vbCrLf
    msgstr = msgstr + "          ��װ   Ԫ������Ϊ   �� " & SmtPartNum & vbCrLf
    msgstr = msgstr + "          ��װ   Ԫ������Ϊ   �� " & LeadPartNum & vbCrLf
    msgstr = msgstr + "          ����   Ԫ������Ϊ   �� " & OtherPartNum & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          None   Ԫ������Ϊ   �� " & NonePartNum & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     Ԫ������Ϊ   �� " & NcPartNum & vbCrLf
    msgstr = msgstr + "          DBG    Ԫ������Ϊ   �� " & DbgPartNum & vbCrLf
    msgstr = msgstr + "          DBG_NC Ԫ������Ϊ   �� " & DbNcPartNum & vbCrLf & vbCrLf
    msgstr = msgstr + " ������ѯ�ļ��Ѿ���ȷ���ɣ���ʹ��ERPϵͳ��ѯ���������" & vbCrLf & vbCrLf
    msgstr = msgstr + "    ע�⣺���ɵ�PCBA_BOM�ļ���Ҫ����޸ĺ�ſɹ����� "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "Ԫ����Ϣ"
    
    xlBook.Close (True) '�رչ�����
    PLxlBook.Close (True)
    NBxlBook.Close (True)
    NonexlBook.Close (True)
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    '����Ԫ������������������
    PartNum(0) = NcPartNum
    PartNum(1) = DbgPartNum
    PartNum(2) = DbNcPartNum
    PartNum(3) = LeadPartNum
    PartNum(4) = SmtPartNum
    PartNum(5) = OtherPartNum
    
    Process 50, "��������ѯ�����������ѡ��tsv�ļ�..."
    Exit Function
    
ErrorHandle:

    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    MsgBox "����BOM�м��ļ�ʱ�����쳣", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
    
End Function

Function xlsInsert(xlSheet As Excel.Worksheet, PartNum As Integer, Row As Long, insertStr() As String)
    If PartNum > 1 Then
        xlSheet.Rows(Row + PartNum & ":" & Row + PartNum).Insert
    End If
    xlSheet.Cells(PartNum + Row, 1) = PartNum
    xlSheet.Cells(PartNum + Row, 2) = insertStr(BomPartNumber)
    xlSheet.Cells(PartNum + Row, 8) = insertStr(BomValue)
    xlSheet.Cells(PartNum + Row, 5) = insertStr(BomQuantity)
    xlSheet.Cells(PartNum + Row, 6) = insertStr(BomPartRef)
    xlSheet.Cells(PartNum + Row, 7) = insertStr(BomPCBfootprint)
End Function

Function ExcelCreate()

    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook, PLxlBook As Excel.Workbook, NBxlBook As Excel.Workbook, NonexlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet, PLxlSheet As Excel.Worksheet, NBxlSheet As Excel.Worksheet, NonexlSheet As Excel.Worksheet
    
    Dim rngNC           As Range
    Dim rngDB           As Range
    Dim rngDBNC         As Range
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    'PCBA_BOM
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
    Set xlSheet = xlBook.Worksheets(1)
    xlBook.SaveAs (SaveAsPath & "_PCBA_BOM.xls")
    xlBook.Close (True) '�رչ�����
    
    '����������Դ��ѯxls
    Set PLxlBook = xlApp.Workbooks.Open(App.Path & "\template\������ѯ_template.xls")
    Set PLxlSheet = PLxlBook.Worksheets(1)
    PLxlBook.SaveAs (SaveAsPath & "_������Դ��ѯ.xls")
    PLxlBook.Close (True)
    

    '����NC_DBGԪ��xls
    Set NBxlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
    Set NBxlSheet = NBxlBook.Worksheets(1)
    NBxlBook.SaveAs (SaveAsPath & "_NC_DBG.xls")
    
    NBxlBook.Close (True)
    
    Set NonexlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
    NonexlBook.SaveAs (SaveAsPath & "_None_PartRef.xls")
    NonexlBook.Worksheets(1).Name = "NoneԪ��"
    
    Set NonexlSheet = NonexlBook.Worksheets(1)
    
    With NonexlSheet.Cells
        Set rngNC = .Find("NCԪ��", lookin:=xlValues)
        Set rngDB = .Find("DBGԪ��", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NCԪ��", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Then
            MsgBox "NC_DBGģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            End
        End If
    End With
    
    '����NoneShleet
    NonexlSheet.Cells(rngNC.Row, 2) = "None"
    NonexlSheet.Rows(rngDB.Row & ":" & rngDBNC.Row + 1).Delete

    NonexlBook.Close (True)
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    Exit Function
    
ErrorHandle:
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    MsgBox "����BOM�м��ļ�ʱ�����쳣", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
    

End Function

