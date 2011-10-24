Attribute VB_Name = "mMakeBOM"
Option Explicit

'BOM����
Public Enum BomType

BOM_ALL = 0
BOM_NCDBG
BOM_NONE

BOM_Ԥ
BOM_����
BOM_����
BOM_����

End Enum

'BMF�ļ������ʽ
'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
'0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--
'BMF_ItemNum=0
'BMF_PartNum
'BMF_Value
'BMF_Quantity
'BMF_PartRef
'BMF_PcbFB
'BMF_MountType
'BMF_Description
'BMF_TP1
'BMF_TP2
'BMF_TP3�

Function xlsInsert(xlSheet As Excel.Worksheet, ItemNum As Integer, Row As Long, insertStr() As String, OrgEnable As Boolean)
    
    '���в���Ҫ����
    If ItemNum > 1 Then
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Insert
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Interior.Pattern = xlNone 'ȥ����ɫ�ȸ�ʽ ������ʾbug
    End If
    
    xlSheet.Cells(ItemNum + Row, 1) = ItemNum
    xlSheet.Cells(ItemNum + Row, 2) = insertStr(BMF_PartNum)
    xlSheet.Cells(ItemNum + Row, 3) = insertStr(BMF_Description)
    xlSheet.Cells(ItemNum + Row, 5) = insertStr(BMF_Quantity)
    xlSheet.Cells(ItemNum + Row, 6) = insertStr(BMF_PartRef)
    xlSheet.Cells(ItemNum + Row, 7) = insertStr(BMF_PcbFB)
    xlSheet.Cells(ItemNum + Row, 8) = insertStr(BMF_Value)
    
    '�Ƿ���ӿ����Ϣ��
    If OrgEnable = True Then
        If insertStr(BMF_TP1) = "-" Then
            xlSheet.Cells(ItemNum + Row, 9) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 9) = insertStr(BMF_TP1)
            If insertStr(BMF_TP1) = "0" Or InStr(insertStr(BMF_TP1), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 9).Interior.Color = 52479 '��ǿ������ɫ��ʾ
            End If
        End If
        
        If insertStr(BMF_TP2) = "-" Then
            xlSheet.Cells(ItemNum + Row, 10) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 10) = insertStr(BMF_TP2)
            If insertStr(BMF_TP2) = "0" Or InStr(insertStr(BMF_TP2), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 10).Interior.Color = 52479 '��ǿ������ɫ��ʾ
            End If
        End If
            
        If insertStr(BMF_TP3) = "-" Then
            xlSheet.Cells(ItemNum + Row, 11) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 11) = insertStr(BMF_TP3)
            If insertStr(BMF_TP3) = "0" Or InStr(insertStr(BMF_TP3), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 11).Interior.Color = 52479 '��ǿ������ɫ��ʾ
            End If
        End If
    End If
    
End Function

Function CopyLine(xlSheetTo As Excel.Worksheet, RowTo As Integer, xlSheetFrom As Excel.Worksheet, RowFrom As Integer, ColumnNum As Integer, PartNum As Integer)
    xlSheetTo.Rows(RowTo & ":" & RowTo).Insert
    xlSheetTo.Cells(RowTo, 1) = PartNum
    Dim i As Integer
    For i = 2 To ColumnNum
        xlSheetTo.Cells(RowTo, i) = xlSheetFrom.Cells(RowFrom, i)
    Next i
    xlSheetTo.Rows(RowTo & ":" & RowTo).Font.ColorIndex = 5
End Function

'����ѡ���BOM ��������ʽ Ϊ���������׼��
Function ExcelCreate(bt_value As BomType)

    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    '=====================================================================================
    'ԤBOM��Capture�е�����BOM��NoneԪ����NCԪ����DBGԪ����DBG_NCԪ��֮�������Ԫ���ļ��ϡ�
    '=====================================================================================
    If bt_value = BOM_Ԥ Then
        'PCBA_BOM
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_ԤBOM_BMF.xls")
        
        xlBook.Close (True) '�رչ�����
    End If
    
    '=====================================================================================
    'NC_DBGԪ��xls
    '=====================================================================================
    If bt_value = BOM_NCDBG Then
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_NC_DBG.xls")
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'NoneԪ��xls
    '=====================================================================================
    If bt_value = BOM_NONE Then
    
        Dim rngNC           As Range
        Dim rngDB           As Range
        Dim rngDBNC         As Range
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        xlBook.SaveAs (SaveAsPath & "_None_PartRef.xls")
        xlBook.Worksheets(1).Name = "NoneԪ��"
        
        Set xlSheet = xlBook.Worksheets(1)
        
        With xlSheet.Cells
            Set rngNC = .Find("NCԪ��", lookin:=xlValues)
            Set rngDB = .Find("DBGԪ��", lookin:=xlValues)
            Set rngDBNC = .Find("DBG_NCԪ��", lookin:=xlValues)
            If rngNC Is Nothing Or rngDB Is Nothing Then
                MsgBox "NC_DBGģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
                End
            End If
        End With
        
        '����NoneShleet
        xlSheet.Cells(rngNC.Row, 2) = "None"
        xlSheet.Rows(rngDB.Row & ":" & rngDBNC.Row + 1).Delete
    
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    '����BOM �� ԤBOM + DBGԪ�� - �´�������+ ���Ͽ����Ϣ��
    '=====================================================================================
    If bt_value = BOM_���� Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_����BOM.xls")
         
        '��������BOM��ʽ
        '������BOM�в�����(I��TP1���) (J��TP2���) (K��TP3���)����ѡ������Ϣ��
        '���������Ӧ����ӵ�����
        xlSheet.Columns("C:C").ColumnWidth = 45
        xlSheet.Columns("G:G").ColumnWidth = 12
        xlSheet.Columns("H:H").ColumnWidth = 12
        
        xlSheet.Columns("H:H").Copy
        xlSheet.Columns("I:I").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
        xlSheet.Columns("I:I").Copy
        xlSheet.Columns("J:J").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
        xlSheet.Columns("K:K").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
        xlApp.CutCopyMode = False
        
        '�������
        xlSheet.Cells(5, 9) = "TP1���"
        xlSheet.Cells(5, 10) = "TP2���"
        xlSheet.Cells(5, 11) = "TP3���"

        xlSheet.Cells(5, 1).Select
         
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    '����BOM �� ԤBOM + DBGԪ�� ���ϡ�
    '=====================================================================================
    If bt_value = BOM_���� Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlSheet.SaveAs (SaveAsPath & "_����BOM.xls")
        
        '������ӡ��ʽ
        With xlSheet.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = 80
        End With
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    '����BOM �� ԤBOM + DBG_NCԪ�� ����
    '=====================================================================================
    If bt_value = BOM_���� Then

        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_����BOM.xls")
        
        xlBook.Close (True)
    End If
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    Exit Function
    
ErrorHandle:
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    MsgBox "����BOM�м��ļ�ʱ�����쳣", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
    

End Function

'�����Ӧ������

'=====================================================================================
'1.ԤBOM��Capture�е�����BOM��NoneԪ����NCԪ����DBGԪ����DBG_NCԪ��֮�������Ԫ���ļ��ϡ�
'2.NC_DBGԪ��xls
'3.NoneԪ��xls
'4.����BOM �� ԤBOM + DBGԪ�� - �´�������+ ���Ͽ����Ϣ
'5.����BOM �� ԤBOM + DBGԪ�� ����
'6.����BOM �� ԤBOM + DBG_NCԪ�� ����
'�
'Ԫ������ �� ��ͨԪ�� NcDbgԪ�� NoneԪ�� �������� �
'=====================================================================================
Function CreateBOM(bt_value As BomType) As Boolean
    
    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
     
    '����Ӧ���ļ�
    Select Case bt_value
    Case BOM_NCDBG:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Case BOM_NONE:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
    Case BOM_Ԥ:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_ԤBOM_BMF.xls")
    Case BOM_����:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_����BOM.xls")
    Case BOM_����:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_����BOM.xls")
    Case BOM_����:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_����BOM.xls")
    Case Else
        GoTo ErrorHandle
    End Select
     
    Set xlSheet = xlBook.Worksheets(1)
    
    '��λ����Ԫ��λ��
    Dim rngPos1       As Range '��BOM�ļ��б�ʾ"SMTԪ��"λ�� ��DBG��NoneԪ������"NCԪ��"λ��
    Dim rngPos2       As Range '��BOM�ļ��б�ʾ"DIPԪ��"λ�� ��DBG��NoneԪ������"DBGԪ��"λ��
    Dim rngPos3       As Range '��BOM�ļ��б�ʾ"����Ԫ��"λ�� ��DBG��NoneԪ������"DBG_NCԪ��"λ��
    With xlSheet.Cells
        
        Select Case bt_value
        Case BOM_NCDBG:
            Set rngPos1 = .Find("NCԪ��", lookin:=xlValues)
            Set rngPos2 = .Find("DBGԪ��", lookin:=xlValues)
            Set rngPos3 = .Find("DBG_NCԪ��", lookin:=xlValues)
            
        Case BOM_NONE:
            Set rngPos1 = .Find("None", lookin:=xlValues)
            Set rngPos2 = .Find("None", lookin:=xlValues)
            Set rngPos3 = .Find("None", lookin:=xlValues)
    
        Case BOM_Ԥ, BOM_����, BOM_����, BOM_����:
            Set rngPos1 = .Find("SMTԪ��", lookin:=xlValues)
            Set rngPos2 = .Find("DIPԪ��", lookin:=xlValues)
            Set rngPos3 = .Find("����Ԫ��", lookin:=xlValues)
            
        Case Else
            GoTo ErrorHandle
        End Select
        
        If rngPos1 Is Nothing Or rngPos2 Is Nothing Or rngPos3 Is Nothing Then
            MsgBox "ģ�����-Ԫ��λ�ö�λ����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    
    '=======================================================
    '��ʼ��ȡbmf�ļ��������Ӧ������
    '=======================================================
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    Dim bmfAtom()       As String
    
    Dim ItemNum1        As Integer 'NCԪ�� �� NoneԪ�� �� SMTԪ����
    Dim ItemNum2        As Integer 'DBGԪ�� �� NoneԪ�� �� DIPԪ����
    Dim ItemNum3        As Integer 'DBG_NCԪ�� �� NoneԪ�� �� ����Ԫ����
    
    Dim i               As Integer
    Dim OrgEnable       As Boolean
    
    ItemNum1 = 0
    ItemNum2 = 0
    ItemNum3 = 0
    
    '�Ƿ���ӿ����Ϣ��
    '��Ҫ���tsv�ļ�����ʱ�� ���ʱ�䲻��3���� �����Ϣ�����Ҳ��û���õ�
    If bt_value = BOM_���� Then
        OrgEnable = True
    Else
        OrgEnable = False
    End If
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '����bmf�ļ�
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        
        If bmfAtom(BMF_PcbFB) = "" Or bmfAtom(BMF_Value) = "" Then
                MsgBox "��" & bmfAtom(BMF_ItemNum) & "��Ԫ����Ϣ������", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "����"
                GoTo ErrorHandle
        End If
        
        If InStr(bmfAtom(BMF_Value), "_DBG_NC") > 0 Or bmfAtom(BMF_Value) = "DBG_NC" Then
            If bt_value = BOM_NCDBG Then
                'DBG_NCԪ��
                ItemNum3 = ItemNum3 + 1
                xlsInsert xlSheet, ItemNum3, rngPos3.Row, bmfAtom, OrgEnable
            End If
            
        ElseIf InStr(bmfAtom(BMF_Value), "_DBG") > 0 Or bmfAtom(BMF_Value) = "DBG" Then
            If bt_value = BOM_NCDBG Then
                'DBGԪ��
                ItemNum2 = ItemNum2 + 1
                xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            End If
           
        ElseIf InStr(bmfAtom(BMF_Value), "_NC") > 0 Or bmfAtom(BMF_Value) = "NC" Then
            If bt_value = BOM_NCDBG Then
                'NCԪ��
                ItemNum1 = ItemNum1 + 1
                xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
            End If
            
        Else
            If bt_value = BOM_NONE Then
                If bmfAtom(BMF_MountType) = "N" Then
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                End If
            End If
            
            If bt_value = BOM_���� Or BOM_���� Or BOM_���� Or BOM_Ԥ Then
                '========================================================
                '��ͨԪ�� ����Ԫ����װ���� �Ȳ��������ĸ�BOM ��������
                Select Case bmfAtom(BMF_MountType)
                Case "S":
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                    
                Case "S+":
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                    xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Interior.Color = 16737792
                    
                Case "L":
                    ItemNum2 = ItemNum2 + 1
                    xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
                    
                Case "N":
                    'Do Nothing
                    
                Case Else
                    '���ļ���û�в鵽��װ���ܾ�����BOM
                    MsgBox "δ֪��װ[" & bmfAtom(BMF_PcbFB) & "]������¿��ļ���"
                    GoTo ErrorHandle
                End Select
            End If
            
        End If
        
    Next i
    
    If bt_value = BOM_���� Or BOM_���� Or BOM_���� Or BOM_Ԥ Then
        '�޸Ļ�������
        xlSheet.Cells(2, 1) = "���ͣ�  " & MainForm.ItemNameText.Text & "            PCBA �汾��                       ���Ʒ��ţ�"
        If MainForm.ItemNameText.Text = "" Then
            xlSheet.Cells(2, 1).Font.ColorIndex = 5
        End If
        
    End If
    
    '�������ָ��ֲ�ͬ��BOM
    'ԤBOM��Capture�е�����BOM��NoneԪ����NCԪ����DBGԪ����DBG_NCԪ��֮�������Ԫ���ļ���
    '����BOM �� ԤBOM + DBGԪ�� - �´�������+ ���Ͽ����Ϣ
    '����BOM �� ԤBOM + DBGԪ�� ����
    '����BOM �� ԤBOM + DBG_NCԪ�� ����
    'Select Case bt_value
    
    'End Select
     
    xlBook.Close (True) '�رչ�����

    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    Exit Function
    
ErrorHandle:

    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    MsgBox "����BOM�м��ļ�ʱ�����쳣", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
    
End Function


Function BomAdjust() As Boolean
    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim ����BOM_xlBook As Excel.Workbook
    Dim ����BOM_xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application") '����EXCEL����
    xlApp.Visible = False  '����EXCEL����ɼ����򲻿ɼ���
    
    Set ����BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_����BOM.xls")
    Set ����BOM_xlSheet = ����BOM_xlBook.Worksheets(1)
    
    '������BOM�в�����(I��TP1���) (J��TP2���) (K��TP3���)����ѡ������Ϣ��
    '���������Ӧ����ӵ�����
    ����BOM_xlSheet.Columns("C:C").ColumnWidth = 45
    ����BOM_xlSheet.Columns("G:G").ColumnWidth = 12
    ����BOM_xlSheet.Columns("H:H").ColumnWidth = 12
    
    ����BOM_xlSheet.Columns("H:H").Copy
    ����BOM_xlSheet.Columns("I:I").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    ����BOM_xlSheet.Columns("I:I").Copy
    ����BOM_xlSheet.Columns("J:J").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    ����BOM_xlSheet.Columns("K:K").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    xlApp.CutCopyMode = False
    
    '�������
    ����BOM_xlSheet.Cells(5, 9) = "TP1���"
    ����BOM_xlSheet.Cells(5, 10) = "TP2���"
    ����BOM_xlSheet.Cells(5, 11) = "TP3���"

    ����BOM_xlSheet.Cells(5, 9).Select
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    
    '��ȡԪ������
    LeadPartNum = PartNum(3)
    SmtPartNum = PartNum(4)
    OtherPartNum = PartNum(5)
    
    'ɾ������Ԫ��
    If DelSamplePart(����BOM_xlSheet) = False Then
        GoTo ErrorHandle
    End If
    
    ����BOM_xlBook.Save
    
    '���±��
    If ReNum(����BOM_xlSheet) = False Then
        GoTo ErrorHandle
    End If
        
    ����BOM_xlBook.Save
    ����BOM_xlBook.Close (True) '�رչ�����
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    BomAdjust = True
    Exit Function

ErrorHandle:
    
    ����BOM_xlBook.Close (True) '�رչ�����
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    BomAdjust = False
    
End Function

Function DelSamplePart(xlSheet As Excel.Worksheet) As Boolean
    Dim rngStart        As Range
    Dim rngEND          As Range
    
    'ɾ���������� �ϺŴ��� ��������12345xxxxx ��xxxxx xxxxx����
    With xlSheet.Cells
        Set rngStart = .Find("SMTԪ��", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngStart Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOMģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            DelSamplePart = False
        End If
    End With
    
    Dim i          As Integer
    Dim j          As Integer
    Dim PartNum    As Integer
    Dim DelRows()  As Integer
    
    PartNum = rngEND.Row - rngStart.Row
    ReDim DelRows(PartNum) As Integer
    
    j = 0
    For i = rngStart.Row To rngEND.Row
        If IsNumeric(xlSheet.Cells(i, 2)) = False _
             And xlSheet.Cells(i, 2) <> "SMTԪ��" _
             And xlSheet.Cells(i, 2) <> "DIPԪ��" _
             And xlSheet.Cells(i, 2) <> "����Ԫ��" _
             And xlSheet.Cells(i, 2) <> "END" Then
            DelRows(j) = i
            j = j + 1
        End If
    Next i
    
    For i = 0 To j
        If DelRows(i) <> 0 Then
            xlSheet.Rows(DelRows(i) - i & ":" & DelRows(i) - i).Delete
        End If
    Next i
    
    DelSamplePart = True
    
End Function

Function ReNum(xlSheet As Excel.Worksheet) As Boolean
    '���±��
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    Dim rngEND          As Range
    
    'ɾ���������� �ϺŴ��� ��������12345xxxxx ��xxxxx xxxxx����
    With xlSheet.Cells
        Set rngSMT = .Find("SMTԪ��", lookin:=xlValues)
        Set rngLEAD = .Find("DIPԪ��", lookin:=xlValues)
        Set rngOther = .Find("����Ԫ��", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOMģ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            ReNum = False
        End If
    End With
    
    Dim j As Integer
    For j = 1 To rngLEAD.Row - rngSMT.Row - 1
        xlSheet.Cells(rngSMT.Row + j, 1) = j
    Next j
    
    For j = 1 To rngOther.Row - rngLEAD.Row - 1
        xlSheet.Cells(rngLEAD.Row + j, 1) = j
    Next j
    
    For j = 1 To rngEND.Row - rngOther.Row - 1
        xlSheet.Cells(rngOther.Row + j, 1) = j
    Next j
    
    ReNum = True
    
End Function



