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
Public BmfFilePath     As String  'ԭʼ�ļ�������
Public SaveAsPath      As String  'BOM������ļ�·��
Public tsvFilePath     As String  'tsv�ļ�·����Ϣ

'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
'0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--

'BMF�ļ���Ϣ�����ʽ
Public Enum BmfInfoFormat

BMF_ItemNum = 0
BMF_PartNum
BMF_Value
BMF_Quantity
BMF_PartRef
BMF_PcbFB
BMF_MountType
BMF_Description
BMF_TP1
BMF_TP2
BMF_TP3

BMF_TOTAL = 10

End Enum

Function BuildProjectPath(srcPath As String)
    '��������������Ҫ��Ŀ¼��Ϣ�������������У����˿�д����Щ·��
    Dim tmpPath As String
    
    BomFilePath = srcPath
    ProjectDir = Left(BomFilePath, InStrRev(BomFilePath, "\") - 1) & "\"
    
    If MainForm.ItemNameText.Text <> "" Then
        SaveAsPath = ProjectDir & "BOM\" & MainForm.ItemNameText.Text
    Else
        tmpPath = Right(BomFilePath, Len(BomFilePath) - InStrRev(BomFilePath, "\"))
        tmpPath = ProjectDir & "BOM\" & tmpPath
        SaveAsPath = Left(tmpPath, InStrRev(tmpPath, ".") - 1)
    End If
    
    BmfFilePath = SaveAsPath & ".bmf"
    
    '�ڹ���Ŀ¼�´���BOMĿ¼
    If Dir(ProjectDir & "BOM\", vbDirectory) = "" Then
        MkDir ProjectDir & "BOM\"
    End If
    
    SaveSetting App.EXEName, "ProjectDir", "�ϴι���Ŀ¼", ProjectDir
End Function

Function ClearPath()
    '��������������Ҫ��Ŀ¼��Ϣ�������������У����˿�д����Щ·��
    BomFilePath = ""
    BmfFilePath = ""
    SaveAsPath = ""
    tsvFilePath = ""
End Function

Function BomMakePLExcel()
    
    Dim Bom                As String
    Dim BomLine()          As String
    
    Dim Atom()             As String
    
    Dim PartNum       As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    '����������ѯExcel�ļ�
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\������ѯ_template.xls")
    Set xlSheet = xlBook.Worksheets(1)
    xlBook.SaveAs (SaveAsPath & "_������Դ��ѯ.xls")
    xlBook.Close (True)
    
    '��������Դ��ѯxls
    Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_������Դ��ѯ.xls")
    Set xlSheet = xlBook.Worksheets(1)
    
    Bom = GetFileContents(BomFilePath)
    
    BomLine = Split(Bom, vbCrLf)
    
    Atom = Split(BomLine(0), vbTab)
    
    
    '��Bom����Ϣ����
    For i = 1 To UBound(BomLine) - 1
        Atom = Split(BomLine(i), vbTab)
        
        '�����Ϻ� Ҫ���������������ѯ��Excel��
        If IsNumeric(Atom(BomPartNumber)) = True And Atom(BomPartNumber) <> "" Then
            xlSheet.Cells(PartNum + 1, 1) = Atom(BomPartNumber)
            PartNum = PartNum + 1
        End If
    Next i
    
    xlBook.Close (True) '�رչ�����
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    Process 20, "������ѯ�ļ��������..."
    
    Exit Function
    
ErrorHandle:

    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    MsgBox "����BOM�м��ļ�ʱ�����쳣", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
    
End Function

Function ReadBomFile() As Boolean
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim newbomstr()     As String
    
    Dim i               As Integer
    
    FileContents = GetFileContents(BomFilePath)
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

'������Bom ���BMF�ļ���ʽ�Ƿ���ȷ
Function CheckBmf() As Boolean
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    Dim i               As Integer
    
    '��ʼ������ֵ
    CheckBmf = False
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '������Bom ���BMF�ļ���ʽ�Ƿ���ȷ
    For i = 0 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        If UBound(bmfAtom) = BMF_TOTAL Then
            CheckBmf = True
        End If
    Next i
    
End Function

'���ݸ������ַ������Ҹ������У����ظ����кŵ��ַ���
Function LookupBmfAtom(checkStr As String, checkCol As Integer, returnCol As Integer) As String
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    Dim i               As Integer
    
    '��ʼ������ֵ
    LookupBmfAtom = "-"
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '������Bom ���Ҷ�Ӧ���ַ���
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        If checkStr = bmfAtom(checkCol) Then
            LookupBmfAtom = bmfAtom(returnCol)
        End If
    Next i
    
End Function

'���ݸ������ַ������Ҹ������У����ز��ҵ��ĵ�һ���к�
Function LookupBmfRow(checkStr As String, checkCol As Integer) As Integer
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    Dim i               As Integer
    
    '��ʼ������ֵ
    LookupBmfRow = -1
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '������Bom ���Ҷ�Ӧ���ַ���
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        If checkStr = bmfAtom(checkCol) Then
            LookupBmfRow = i
            Exit For
        End If
    Next i
    
End Function

'���ݸ����кţ������кţ������ַ���
Function GetBmfAtom(Row As Integer, Col As Integer) As String
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String

    
    '��ʼ������ֵ
    GetBmfAtom = ""
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)
    
    bmfAtom = Split(bmfBomLine(Row), vbTab)
    GetBmfAtom = bmfAtom(Col)
    
End Function

'���ݸ����кţ�����һ������
Function GetBmfLine(Row As Integer) As String
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    '��ʼ������ֵ
    GetBmfLine = ""
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    GetBmfLine = bmfBomLine(Row)
    
End Function

'�޸ĸ����кţ������кŵĶ�Ӧ������
Function SetBmfAtom(Row As Integer, Col As Integer, addStr As String)
             
    Dim oldBom          As String
    Dim newBomLine()    As String
    
    Dim BomAtom()       As String
    
    Dim i               As Integer
    
    oldBom = GetFileContents(BmfFilePath)
    
    newBomLine = Split(oldBom, vbCrLf)
    
    '�����ͷ��Ϣ
    'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
    '0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--
    BomAtom = Split(newBomLine(Row), vbTab)
    BomAtom(Col) = addStr
    
    newBomLine(Row) = ""
    
    '����Bom����Ϣ���뵽��Bom ������Bom
    For i = 0 To UBound(BomAtom) - 1
        newBomLine(Row) = newBomLine(Row) + BomAtom(i) + vbTab
    Next i
    
    '���һ��û��vbTab
    newBomLine(Row) = newBomLine(Row) + BomAtom(UBound(BomAtom))
    
    If Dir(BmfFilePath) <> "" Then
        Kill BmfFilePath
    End If
    
    Open BmfFilePath For Binary Access Write As #1
    Seek #1, 1
    Put #1, , newBomLine(0) & vbCrLf
    
    For i = 1 To UBound(newBomLine) - 1
        Put #1, , newBomLine(i) & vbCrLf
    Next i
        
    Put #1, , vbCrLf

    Close #1
    
End Function

Function BmfMaker()
    

    '��ȡ����Ϣ
    Process 10, "��ȡ���ļ���Ϣ..."
    
    Dim leadLibInfo()      As String
    Dim smtLibInfo()       As String
    Dim IgLibInfo()        As String
    
    leadLibInfo = ReadLibs(LIB_LEAD)
    smtLibInfo = ReadLibs(LIB_SMD)
    IgLibInfo = ReadLibs(LIB_NONE)
        
    '����BOM�м��ļ�
    Dim oldBom          As String
    Dim oldBomLine()    As String
    Dim newBomLine()    As String
    
    Dim oldAtom()       As String
    Dim newAtom()       As String
    
    Dim strtmp          As String
    
    Dim NcPartNum       As Integer
    Dim DbgPartNum      As Integer
    Dim DbNcPartNum     As Integer
    Dim NonePartNum     As Integer
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    
    Dim PLPartNum       As Integer
    
    Dim IsLead          As Integer
    Dim IsSmt           As Integer
    Dim IsNone          As Integer
    
    Dim i               As Integer
    Dim j               As Integer
    
    Dim BmfExistFlag    As Boolean
    
    'BMF�ļ��Ƿ���� �Ƿ��������
    BmfExistFlag = CheckBmf
    
    oldBom = GetFileContents(BomFilePath)
    
    oldBomLine = Split(oldBom, vbCrLf)
    newBomLine = Split(oldBom, vbCrLf)
    
    For i = 0 To UBound(oldBomLine)
        newBomLine(i) = ""
    Next i
    
    '�����ͷ��Ϣ
    'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
    '0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--
    oldAtom = Split(oldBomLine(0), vbTab)
    newBomLine(0) = oldAtom(BomItemNumber) + vbTab + oldAtom(BomPartNumber) + vbTab
    newBomLine(0) = newBomLine(0) + oldAtom(BomValue) + vbTab + oldAtom(BomQuantity) + vbTab
    newBomLine(0) = newBomLine(0) + oldAtom(BomPartRef) + vbTab + oldAtom(BomPCBfootprint) + vbTab
    '��װ��ʽ��Ϣ������������Ϣ��
    newBomLine(0) = newBomLine(0) + "Mount Type" + vbTab + "Description" + vbTab
    '�����Ϣ
    newBomLine(0) = newBomLine(0) + "TP1" + vbTab + "TP2" + vbTab + "TP3"
    
    
    '����Bom����Ϣ���뵽��Bom ������Bom
    'On Error GoTo ErrorHandle
    For i = 1 To UBound(oldBomLine) - 1
        oldAtom = Split(oldBomLine(i), vbTab)
        
         '�Ϻź�������������ϻ���
        If UBound(oldAtom) < 5 Then
            strtmp = oldBomLine(i)
            For j = i + 1 To UBound(oldBomLine)
                If Len(oldBomLine(j)) > 1 Then
                    oldAtom = Split(strtmp & oldBomLine(j), vbTab)
                    i = j
                    Exit For
                End If
            Next
        End If
        
        newBomLine(i) = oldAtom(BomItemNumber) + vbTab + oldAtom(BomPartNumber) + vbTab
        newBomLine(i) = newBomLine(i) + oldAtom(BomValue) + vbTab + oldAtom(BomQuantity) + vbTab
        newBomLine(i) = newBomLine(i) + oldAtom(BomPartRef) + vbTab + oldAtom(BomPCBfootprint) + vbTab
        
        Process i * 40 / UBound(oldBomLine) + 10, "������װ[" & oldAtom(BomPCBfootprint) & "]..."
        
        '�����Ϻ�
        If IsNumeric(oldAtom(BomPartNumber)) = True And oldAtom(BomPartNumber) <> "" Then
            PLPartNum = PLPartNum + 1
        End If
        
        'ͳ��Ԫ�����͸���
        If InStr(oldAtom(BomValue), "_DBG_NC") > 0 Or oldAtom(BomValue) = "DBG_NC" Then
            'DBG_NCԪ��
            DbNcPartNum = DbNcPartNum + 1
            
        ElseIf InStr(oldAtom(BomValue), "_DBG") > 0 Or oldAtom(BomValue) = "DBG" Then
            'DBGԪ��
            DbgPartNum = DbgPartNum + 1
           
        ElseIf InStr(oldAtom(BomValue), "_NC") > 0 Or oldAtom(BomValue) = "NC" Then
            'NCԪ��
            NcPartNum = NcPartNum + 1
            
        End If
        
        
        If oldAtom(BomPCBfootprint) = "" Then
            MsgBox oldAtom(BomPartNumber) & "��PCB footprintΪ��", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "����"
            Exit Function
        End If
        
        '========================================================
        '����Ԫ����װ����  ���Mount Type��
        
        '�ж�Ԫ������
        IsLead = QueryLib(leadLibInfo, oldAtom(BomPCBfootprint))
        IsSmt = QueryLib(smtLibInfo, oldAtom(BomPCBfootprint))
        IsNone = QueryLib(IgLibInfo, oldAtom(BomPCBfootprint))
            
        If IsLead = 1 And IsSmt = 0 And IsNone = 0 Then
            'ͳ�Ʋ�д��LEADԪ��
            LeadPartNum = LeadPartNum + 1
            newBomLine(i) = newBomLine(i) + "L" + vbTab
            
        ElseIf IsLead = 0 And IsSmt = 1 And IsNone = 0 Then
            'ͳ�Ʋ�д��SMTԪ��
            SmtPartNum = SmtPartNum + 1
            newBomLine(i) = newBomLine(i) + "S" + vbTab
        
        ElseIf IsLead = 0 And IsSmt = 0 And IsNone = 1 Then
            'ͳ�Ʋ�д�뵥�����ļ��� NoneԪ��
            NonePartNum = NonePartNum + 1
            newBomLine(i) = newBomLine(i) + "N" + vbTab
            
        ElseIf IsLead = 1 And IsSmt = 1 And IsNone = 0 Then
            '����SMTԪ�� �������� ������ɫ��ʾ
            SmtPartNum = SmtPartNum + 1
            newBomLine(i) = newBomLine(i) + "S+" + vbTab
            
        Else
           '���ļ���û�в鵽��װ���ܾ�����BOM
            MsgBox "��װ[" & oldAtom(BomPCBfootprint) & "]�������ڿ��ļ��У�����¿��ļ���"
            Exit Function
            
        End If
        
        IsLead = 0
        IsSmt = 0
        IsNone = 0
        
        '���Description TPn����Ϣ ������ھɵ�.bmf�ļ������Դ��䵼�룬��ԼЧ��
        '��װ��ʽ��Ϣ������������Ϣ��
        If BmfExistFlag Then
            '���ھ��ļ����������Ϣ ���Ա��Ϻ� �ϺŲ���� �޷�����
            If IsNumeric(oldAtom(BomPartNumber)) = True And oldAtom(BomPartNumber) <> "" Then
                'BomPartNumber��keyed
                '������Ϣ
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_Description) + vbTab
                '�����Ϣ
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_TP1) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_TP2) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_TP3)
            
            ElseIf oldAtom(BomValue) <> "" Then
                'BomValueҲ��Keyed
                '������Ϣ
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_Description) + vbTab
                '�����Ϣ
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_TP1) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_TP2) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_TP3)
            Else
                '�������Ӧ����
                newBomLine(i) = newBomLine(i) + "-" + vbTab
                newBomLine(i) = newBomLine(i) + "-" + vbTab + "-" + vbTab + "-"
            End If
        Else
            '�������Ӧ����
            newBomLine(i) = newBomLine(i) + "-" + vbTab
            newBomLine(i) = newBomLine(i) + "-" + vbTab + "-" + vbTab + "-"
        End If
        
    Next i
    
    
    If Dir(BmfFilePath) <> "" Then
        Kill BmfFilePath
    End If
    
    Open BmfFilePath For Binary Access Write As #1
    Seek #1, 1
    Put #1, , newBomLine(0) & vbCrLf
    
    For j = 1 To UBound(newBomLine) - 1
        Put #1, , newBomLine(j) & vbCrLf
    Next j
        
    'Put #1, , vbCrLf

    Close #1
    
    '����Ԫ������������������
    PartNum(0) = NcPartNum
    PartNum(1) = DbgPartNum
    PartNum(2) = DbNcPartNum
    PartNum(3) = LeadPartNum
    PartNum(4) = SmtPartNum
    PartNum(5) = 0
    
    Process 50, "BOM�м��ļ����ɳɹ�..."
    
End Function

Function ImportTSV() As Boolean

    Process 51, "����tsv�ļ���Ϣ..."
    
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim bomstr()        As String
    
    Dim i               As Integer
    Dim FindRow         As Integer
    
    '�Զ���Ӧ��ͬ��tsv�ļ�����
    FileContents = UEFLoadTextFile(tsvFilePath, UEF_AUTO)
    fileinfo = Split(FileContents, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������
    
    '��ȡ�������
    Dim SelStorage As String
    
    SelStorage = GetSetting(App.EXEName, "SelectStorage", "�������", "TP1")
        
    
    '���    ����(����)    ״̬    ����    ��λ    �����ϵ    �ܿ� ����   ���� ����
    '0       1             2       3       4       5           6           7
    
    For i = 1 To UBound(fileinfo) - 1
        bomstr = Split(fileinfo(i), vbTab)
        
        Process i * 3 / UBound(fileinfo) + 51 + 1, "��������  [" & bomstr(1) & "]..."
        
        '���Ҳ����������Ϣ
        FindRow = LookupBmfRow(bomstr(1), BMF_PartNum)
        If FindRow > 0 Then
            '��������
            SetBmfAtom FindRow, BMF_Description, bomstr(3)
            
            '�����Ϣ
            Select Case SelStorage
            Case "TP1"
                SetBmfAtom FindRow, BMF_TP1, bomstr(7)
            Case "TP2"
                SetBmfAtom FindRow, BMF_TP2, bomstr(7)
            Case "TP3"
                SetBmfAtom FindRow, BMF_TP3, bomstr(7)
            Case Else
                SetBmfAtom FindRow, BMF_TP1, bomstr(7)
            End Select
            
        End If
        
    Next i
    
    Process 60, "BOM�м��ļ��������..."
 
End Function

Function BmfToAnsi() As Boolean
    
    'ת�������ʽ
    If UEFSaveTextFile(BmfFilePath, UEFLoadTextFile(BmfFilePath, UEF_AUTO), False, UEF_ANSI) = False Then
        MsgBox "bmf�ļ���ʽת������", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
    End If
 
End Function
