Attribute VB_Name = "mMakeBOM"
Option Explicit

Function CreateAllBOM() As Boolean
    
    On Error GoTo ErrorHandle
    
    Process 54, "׼����������BOM������BOM������BOM..."
        
    Dim xlApp As Excel.Application
    Dim PCBA_BOM_xlBook As Excel.Workbook, NCDBBOM_xlBook As Excel.Workbook
    Dim ����BOM_xlBook As Excel.Workbook, ����BOM_xlBook As Excel.Workbook, ����BOM_xlBook As Excel.Workbook
    
    Dim PCBA_BOM_xlSheet As Excel.Worksheet, NCDBBOM_xlSheet As Excel.Worksheet
    Dim ����BOM_xlSheet As Excel.Worksheet, ����BOM_xlSheet As Excel.Worksheet, ����BOM_xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Process 55, "��������BOM..."
    '����BOM
    Set ����BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set ����BOM_xlSheet = ����BOM_xlBook.Worksheets(1)
    ����BOM_xlBook.SaveAs (SaveAsPath & "_����BOM.xls")
    
    
    Process 56, "��������BOM..."
    '����BOM
    Set ����BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set ����BOM_xlSheet = ����BOM_xlBook.Worksheets(1)
    ����BOM_xlBook.SaveAs (SaveAsPath & "_����BOM.xls")
    
    Set NCDBBOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Set NCDBBOM_xlSheet = NCDBBOM_xlBook.Worksheets(1)
    
    Set PCBA_BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set PCBA_BOM_xlSheet = PCBA_BOM_xlBook.Worksheets(1)
    
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    
    Dim rngNC           As Range
    Dim rngDB           As Range
    Dim rngDBNC         As Range
    
    Dim NcPartNum       As Integer
    Dim DbgPartNum      As Integer
    Dim DbNcPartNum     As Integer
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    NcPartNum = PartNum(0)
    DbgPartNum = PartNum(1)
    DbNcPartNum = PartNum(2)
    LeadPartNum = PartNum(3)
    SmtPartNum = PartNum(4)
    OtherPartNum = PartNum(5)
    
    '��λ����Ԫ��λ��
    With NCDBBOM_xlSheet.Cells
        Set rngNC = .Find("NCԪ��", lookin:=xlValues)
        Set rngDB = .Find("DBGԪ��", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NCԪ��", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Then
            MsgBox "NC_DBG�ļ�����", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    '��λ����Ԫ��λ��
    With PCBA_BOM_xlSheet.Cells
        Set rngSMT = .Find("SMTԪ��", lookin:=xlValues)
        Set rngLEAD = .Find("DIPԪ��", lookin:=xlValues)
        Set rngOther = .Find("����Ԫ��", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Then
            MsgBox "PCBA_BOM�������", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            GoTo ErrorHandle
        End If
    End With
    
    '========================================================
    '��ȡ����Ϣ
    'LEAD ��
    Dim leadLibInfo()      As String
    Dim smtLibInfo()       As String
    Dim IgLibInfo()        As String
    leadLibInfo = ReadLibs(LIB_LEAD)
    smtLibInfo = ReadLibs(LIB_SMD)
    IgLibInfo = ReadLibs(LIB_NONE)
    
    Dim i       As Integer
    Dim rngNum  As Range
    
    '��������BOM
    For i = 1 To DbgPartNum
        'MsgBox NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2)
        Process i * 10 / DbgPartNum + 57, "��������---[" & NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2) & "]..."
        
        If NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2) = "" Then
            MsgBox "DBGԪ���ϺŲ����ڣ�NC_DBG_BOM��DBGԪ�����Ϊ" & i, vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
            GoTo ErrorHandle
        Else
            If IsNumeric(NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2)) = True Then
                'MsgBox NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2)
                With ����BOM_xlSheet.Cells
                    Set rngNum = .Find(NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2), lookin:=xlValues)
                    If rngNum Is Nothing Then
                        If QueryLib(smtLibInfo, NCDBBOM_xlSheet.Cells(i + rngDB.Row, 7)) Then
                            SmtPartNum = SmtPartNum + 1
                            CopyLine ����BOM_xlSheet, rngSMT.Row + SmtPartNum, NCDBBOM_xlSheet, i + rngDB.Row, 8, SmtPartNum
                        ElseIf QueryLib(leadLibInfo, NCDBBOM_xlSheet.Cells(i + rngDB.Row, 7)) Then
                            LeadPartNum = LeadPartNum + 1
                            CopyLine ����BOM_xlSheet, rngLEAD.Row + LeadPartNum, NCDBBOM_xlSheet, i + rngDB.Row, 8, LeadPartNum
                        Else
                            OtherPartNum = OtherPartNum + 1
                            CopyLine ����BOM_xlSheet, rngOther.Row + OtherPartNum, NCDBBOM_xlSheet, i + rngDB.Row, 8, OtherPartNum
                        End If
                    Else
                        ����BOM_xlSheet.Cells(rngNum.Row, 5) = CInt(����BOM_xlSheet.Cells(rngNum.Row, 5)) + CInt(NCDBBOM_xlSheet.Cells(i + rngDB.Row, 5))
                        ����BOM_xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
                        ����BOM_xlSheet.Cells(rngNum.Row, 6) = ����BOM_xlSheet.Cells(rngNum.Row, 6) + " " + NCDBBOM_xlSheet.Cells(i + rngDB.Row, 6)
                        ����BOM_xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
                    End If
                End With
            End If
        End If
    Next i
     
     '��������BOM
    For i = 1 To DbNcPartNum
        Process i * 10 / DbNcPartNum + 68, "��������---[" & NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2) & "]..."
    
        If NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2) = "" Then
            MsgBox "DBG_NCԪ���ϺŲ����ڣ�NC_DBG_BOM��DBG_NCԪ�����Ϊ" & i, vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
            GoTo ErrorHandle
        Else
            If IsNumeric(NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2)) = True Then
                With ����BOM_xlSheet.Cells
                    Set rngNum = .Find(NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2), lookin:=xlValues)
                    If rngNum Is Nothing Then
                        If QueryLib(smtLibInfo, NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 7)) Then
                            SmtPartNum = SmtPartNum + 1
                            CopyLine ����BOM_xlSheet, rngSMT.Row + SmtPartNum, NCDBBOM_xlSheet, i + rngDBNC.Row, 8, SmtPartNum
                        ElseIf QueryLib(leadLibInfo, NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 7)) Then
                            LeadPartNum = LeadPartNum + 1
                            CopyLine ����BOM_xlSheet, rngLEAD.Row + LeadPartNum, NCDBBOM_xlSheet, i + rngDBNC.Row, 8, LeadPartNum
                        Else
                            OtherPartNum = OtherPartNum + 1
                            CopyLine ����BOM_xlSheet, rngOther.Row + OtherPartNum, NCDBBOM_xlSheet, i + rngDBNC.Row, 8, OtherPartNum
                        End If
                    Else
                        ����BOM_xlSheet.Cells(rngNum.Row, 5) = CInt(����BOM_xlSheet.Cells(rngNum.Row, 5)) + CInt(NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 5))
                        ����BOM_xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
                        ����BOM_xlSheet.Cells(rngNum.Row, 6) = ����BOM_xlSheet.Cells(rngNum.Row, 6) + " " + NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 6)
                        ����BOM_xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
                    End If
                End With
            End If
        End If
    Next i
    
    Process 78, "��������BOM������BOM..."
    
    ����BOM_xlBook.Save
    ����BOM_xlBook.Save
    ����BOM_xlBook.Close (True) '�رչ�����
    ����BOM_xlBook.Close (True)
    
    Process 79, "���ɵ���BOM..."
    '����BOM��Ҫ������BOM������ʽ�޸ģ����ڴ�ӡ
    Set ����BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_����BOM.xls")
    Set ����BOM_xlSheet = ����BOM_xlBook.Worksheets(1)
    ����BOM_xlBook.SaveAs (SaveAsPath & "_����BOM.xls")
    
    Process 80, "�޸ĵ���BOM�Ĵ�ӡ��ʽ���Ա��ڴ�ӡ..."
    With ����BOM_xlSheet.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = 80
    End With

    Process 81, "�������BOM..."
    
    ����BOM_xlBook.Save
    ����BOM_xlBook.Close (True)
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    CreateAllBOM = True
    Exit Function

ErrorHandle:
    ����BOM_xlBook.Close (True) '�رչ�����
    ����BOM_xlBook.Close (True)
    ����BOM_xlBook.Close (True)
    
    xlApp.Quit '����EXCEL����
    Set xlApp = Nothing '�ͷ�xlApp����
    
    CreateAllBOM = False
    
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

Function ImportTSV(TmpBomFilePath As String, ProcNum As Integer) As Boolean

    Process ProcNum, "����tsv�ļ���Ϣ..."
    
    On Error GoTo ErrorHandle
    
    Dim MSxlApp As Excel.Application
    Dim MSxlBook As Excel.Workbook
    Dim MSxlSheet As Excel.Worksheet
    
    Set MSxlApp = CreateObject("Excel.Application") '����EXCEL����
    Set MSxlBook = MSxlApp.Workbooks.Open(TmpBomFilePath) '���Ѿ����ڵ�BOMģ��
    MSxlApp.Visible = False  '����EXCEL����ɼ����򲻿ɼ���
    
    Set MSxlSheet = MSxlBook.Worksheets(1) '���û������
    
    Dim tsvcode         As String
    Dim tsvDefFmt       As UnicodeEncodeFormat
    
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim bomstr()        As String
    Dim tPartNum        As String
    
    Dim i               As Integer
    Dim j               As Integer
    Dim m               As Integer
    Dim n               As Integer
    Dim rngNum          As Range
    Dim rngZD           As Range
    
    '��Ӧ��ͬ��tsv�ļ�����
    tsvcode = GetSetting(App.EXEName, "tsvEncoder", "tsv�ļ�����", "UTF-8")
    
    Select Case tsvcode
        Case "UTF-8"
            tsvDefFmt = UEF_UTF8
        Case "ANSI"
            tsvDefFmt = UEF_ANSI
        Case "UTF-16LE"
            tsvDefFmt = UEF_UTF16LE
        Case "UTF-16BE"
            tsvDefFmt = UEF_UTF16BE
        Case Else
            tsvDefFmt = UEF_Auto
    End Select
    
    'ת�������ʽ
    If UEFSaveTextFile(tsvFilePath & "_ansi.tsv", UEFLoadTextFile(tsvFilePath, tsvDefFmt), False, UEF_ANSI) = False Then
        MsgBox "tsv�ļ���ȡת������", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
        GoTo ErrorHandle
    End If
        
    FileContents = UEFLoadTextFile(tsvFilePath & "_ansi.tsv", UEF_Auto)
    Kill tsvFilePath & "_ansi.tsv"       'ɾ���м��ļ�
    fileinfo = Split(FileContents, vbCrLf) 'ȡ��Դ�ļ����������ջس��������ָ�������
    
    '��ȡ�������
    Dim SelStorage As String
    Dim StorageNum As Integer
    If InStr(TmpBomFilePath, "����BOM") > 1 Then
        SelStorage = GetSetting(App.EXEName, "SelectStorage", "�������", "TP1")
        
        Select Case SelStorage
        Case "TP1"
            StorageNum = 1
        Case "TP2"
            StorageNum = 2
        Case "TP3"
            StorageNum = 3
        Case Else
            
            StorageNum = 1
        End Select
        
    End If
    
    '���    ����    ״̬    ����    ��λ    �����ϵ    �ܿ� ����   ���� ����
    '0       1       2       3       4       5           6           7
    j = 0
    
    For j = 1 To UBound(fileinfo) - 1
        bomstr = Split(fileinfo(j), vbTab)
        
        Process j * 3 / UBound(fileinfo) + ProcNum + 1, "��������---[" & bomstr(1) & "]..."
        
        '��λ����Ԫ��λ��
        With MSxlSheet.Cells
            Set rngNum = .Find(bomstr(1), lookin:=xlValues)
            If rngNum Is Nothing Then
                'MsgBox ("�Ҳ���" & bomstr(1) & "�ϺŵĶ�Ӧ��������ָ��")
                For m = 1 To Len(bomstr(5))
                    For n = 1 To Len(bomstr(5))
                        tPartNum = Mid(bomstr(5), m, n)
                        If IsNumeric(tPartNum) = True And Len(tPartNum) = Len(bomstr(1)) Then
                            With MSxlSheet.Cells
                                Set rngZD = .Find(tPartNum, lookin:=xlValues)
                                    If rngZD Is Nothing Then
                                        MsgBox "�Ҳ���" & bomstr(1) & "�ϺŵĶ�Ӧ,�����tsv�ļ�", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "����"
                                        GoTo ErrorHandle
                                    Else
                                        'MSxlSheet.Cells(rngZD.Row, 4) = MSxlSheet.Cells(rngZD.Row, 4) & "����" & bomstr(1) & vbCrLf
                                    End If
                            End With
                        End If
                    Next
                Next
            Else
                '�������������
                MSxlSheet.Cells(rngNum.Row, 3) = bomstr(3)
                
                '����BOM����Ҫ��ӽ��ڿ�������˵�� bomstr(7) Ϊ���ڿ�����
                If InStr(TmpBomFilePath, "����BOM") > 1 Then
                    
                    If StorageNum >= 1 And StorageNum <= 3 Then
                        If bomstr(7) = "0" Then
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8).Font.Size = 8
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8).Interior.Color = 52479
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8) = "���ڿ�����Ϊ" & bomstr(7)
                        ElseIf InStr(bomstr(7), "-") = 1 Then
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8).Interior.Color = 52479
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8) = bomstr(7) '���ڿ�����Ϊ��ֵ
                        Else
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8) = bomstr(7) '���ڿ�����
                        End If
                    Else
                        MsgBox "��ѡ���棡"
                        GoTo ErrorHandle
                    End If
                End If
            End If
        End With
    Next j
        
    MSxlBook.Close (True) '�رչ�����
    MSxlApp.Quit '����EXCEL����
    Set MSxlApp = Nothing '�ͷ�xlApp����
    
    ImportTSV = True
    Exit Function

ErrorHandle:
    MSxlBook.Close (True) '�رչ�����
    MSxlApp.Quit '����EXCEL����
    Set MSxlApp = Nothing '�ͷ�xlApp����
    
    ImportTSV = False
    
End Function


