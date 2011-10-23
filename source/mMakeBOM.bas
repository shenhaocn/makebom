Attribute VB_Name = "mMakeBOM"
Option Explicit

Function CreateAllBOM() As Boolean
    
    On Error GoTo ErrorHandle
    
    Process 54, "准备生成领料BOM、调试BOM、生产BOM..."
        
    Dim xlApp As Excel.Application
    Dim PCBA_BOM_xlBook As Excel.Workbook, NCDBBOM_xlBook As Excel.Workbook
    Dim 领料BOM_xlBook As Excel.Workbook, 调试BOM_xlBook As Excel.Workbook, 生产BOM_xlBook As Excel.Workbook
    
    Dim PCBA_BOM_xlSheet As Excel.Worksheet, NCDBBOM_xlSheet As Excel.Worksheet
    Dim 领料BOM_xlSheet As Excel.Worksheet, 调试BOM_xlSheet As Excel.Worksheet, 生产BOM_xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Process 55, "生成领料BOM..."
    '领料BOM
    Set 领料BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set 领料BOM_xlSheet = 领料BOM_xlBook.Worksheets(1)
    领料BOM_xlBook.SaveAs (SaveAsPath & "_领料BOM.xls")
    
    
    Process 56, "生成生产BOM..."
    '生产BOM
    Set 生产BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set 生产BOM_xlSheet = 生产BOM_xlBook.Worksheets(1)
    生产BOM_xlBook.SaveAs (SaveAsPath & "_生产BOM.xls")
    
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
    
    '定位各种元件位置
    With NCDBBOM_xlSheet.Cells
        Set rngNC = .Find("NC元件", lookin:=xlValues)
        Set rngDB = .Find("DBG元件", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NC元件", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Then
            MsgBox "NC_DBG文件错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            GoTo ErrorHandle
        End If
    End With
    
    '定位各种元件位置
    With PCBA_BOM_xlSheet.Cells
        Set rngSMT = .Find("SMT元件", lookin:=xlValues)
        Set rngLEAD = .Find("DIP元件", lookin:=xlValues)
        Set rngOther = .Find("其他元件", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Then
            MsgBox "PCBA_BOM载入错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            GoTo ErrorHandle
        End If
    End With
    
    '========================================================
    '读取库信息
    'LEAD 库
    Dim leadLibInfo()      As String
    Dim smtLibInfo()       As String
    Dim IgLibInfo()        As String
    leadLibInfo = ReadLibs(LIB_LEAD)
    smtLibInfo = ReadLibs(LIB_SMD)
    IgLibInfo = ReadLibs(LIB_NONE)
    
    Dim i       As Integer
    Dim rngNum  As Range
    
    '生成领料BOM
    For i = 1 To DbgPartNum
        'MsgBox NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2)
        Process i * 10 / DbgPartNum + 57, "分析物料---[" & NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2) & "]..."
        
        If NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2) = "" Then
            MsgBox "DBG元件料号不存在，NC_DBG_BOM中DBG元件序号为" & i, vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
            GoTo ErrorHandle
        Else
            If IsNumeric(NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2)) = True Then
                'MsgBox NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2)
                With 领料BOM_xlSheet.Cells
                    Set rngNum = .Find(NCDBBOM_xlSheet.Cells(i + rngDB.Row, 2), lookin:=xlValues)
                    If rngNum Is Nothing Then
                        If QueryLib(smtLibInfo, NCDBBOM_xlSheet.Cells(i + rngDB.Row, 7)) Then
                            SmtPartNum = SmtPartNum + 1
                            CopyLine 领料BOM_xlSheet, rngSMT.Row + SmtPartNum, NCDBBOM_xlSheet, i + rngDB.Row, 8, SmtPartNum
                        ElseIf QueryLib(leadLibInfo, NCDBBOM_xlSheet.Cells(i + rngDB.Row, 7)) Then
                            LeadPartNum = LeadPartNum + 1
                            CopyLine 领料BOM_xlSheet, rngLEAD.Row + LeadPartNum, NCDBBOM_xlSheet, i + rngDB.Row, 8, LeadPartNum
                        Else
                            OtherPartNum = OtherPartNum + 1
                            CopyLine 领料BOM_xlSheet, rngOther.Row + OtherPartNum, NCDBBOM_xlSheet, i + rngDB.Row, 8, OtherPartNum
                        End If
                    Else
                        领料BOM_xlSheet.Cells(rngNum.Row, 5) = CInt(领料BOM_xlSheet.Cells(rngNum.Row, 5)) + CInt(NCDBBOM_xlSheet.Cells(i + rngDB.Row, 5))
                        领料BOM_xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
                        领料BOM_xlSheet.Cells(rngNum.Row, 6) = 领料BOM_xlSheet.Cells(rngNum.Row, 6) + " " + NCDBBOM_xlSheet.Cells(i + rngDB.Row, 6)
                        领料BOM_xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
                    End If
                End With
            End If
        End If
    Next i
     
     '生成生产BOM
    For i = 1 To DbNcPartNum
        Process i * 10 / DbNcPartNum + 68, "分析物料---[" & NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2) & "]..."
    
        If NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2) = "" Then
            MsgBox "DBG_NC元件料号不存在，NC_DBG_BOM中DBG_NC元件序号为" & i, vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
            GoTo ErrorHandle
        Else
            If IsNumeric(NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2)) = True Then
                With 生产BOM_xlSheet.Cells
                    Set rngNum = .Find(NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 2), lookin:=xlValues)
                    If rngNum Is Nothing Then
                        If QueryLib(smtLibInfo, NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 7)) Then
                            SmtPartNum = SmtPartNum + 1
                            CopyLine 生产BOM_xlSheet, rngSMT.Row + SmtPartNum, NCDBBOM_xlSheet, i + rngDBNC.Row, 8, SmtPartNum
                        ElseIf QueryLib(leadLibInfo, NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 7)) Then
                            LeadPartNum = LeadPartNum + 1
                            CopyLine 生产BOM_xlSheet, rngLEAD.Row + LeadPartNum, NCDBBOM_xlSheet, i + rngDBNC.Row, 8, LeadPartNum
                        Else
                            OtherPartNum = OtherPartNum + 1
                            CopyLine 生产BOM_xlSheet, rngOther.Row + OtherPartNum, NCDBBOM_xlSheet, i + rngDBNC.Row, 8, OtherPartNum
                        End If
                    Else
                        生产BOM_xlSheet.Cells(rngNum.Row, 5) = CInt(生产BOM_xlSheet.Cells(rngNum.Row, 5)) + CInt(NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 5))
                        生产BOM_xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
                        生产BOM_xlSheet.Cells(rngNum.Row, 6) = 生产BOM_xlSheet.Cells(rngNum.Row, 6) + " " + NCDBBOM_xlSheet.Cells(i + rngDBNC.Row, 6)
                        生产BOM_xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
                    End If
                End With
            End If
        End If
    Next i
    
    Process 78, "保存领料BOM、生产BOM..."
    
    领料BOM_xlBook.Save
    生产BOM_xlBook.Save
    领料BOM_xlBook.Close (True) '关闭工作簿
    生产BOM_xlBook.Close (True)
    
    Process 79, "生成调试BOM..."
    '调试BOM需要在领料BOM上做格式修改，便于打印
    Set 调试BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_领料BOM.xls")
    Set 调试BOM_xlSheet = 调试BOM_xlBook.Worksheets(1)
    调试BOM_xlBook.SaveAs (SaveAsPath & "_调试BOM.xls")
    
    Process 80, "修改调试BOM的打印格式，以便于打印..."
    With 调试BOM_xlSheet.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = 80
    End With

    Process 81, "保存调试BOM..."
    
    调试BOM_xlBook.Save
    调试BOM_xlBook.Close (True)
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    CreateAllBOM = True
    Exit Function

ErrorHandle:
    领料BOM_xlBook.Close (True) '关闭工作簿
    生产BOM_xlBook.Close (True)
    调试BOM_xlBook.Close (True)
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
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
    Dim 领料BOM_xlBook As Excel.Workbook
    Dim 领料BOM_xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application") '创建EXCEL对象
    xlApp.Visible = False  '设置EXCEL对象可见（或不可见）
    
    Set 领料BOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_领料BOM.xls")
    Set 领料BOM_xlSheet = 领料BOM_xlBook.Worksheets(1)
    
    '在领料BOM中插入列(I：TP1库存) (J：TP2库存) (K：TP3库存)（需选择库存信息）
    '调整表格适应新添加的内容
    领料BOM_xlSheet.Columns("C:C").ColumnWidth = 45
    领料BOM_xlSheet.Columns("G:G").ColumnWidth = 12
    领料BOM_xlSheet.Columns("H:H").ColumnWidth = 12
    
    领料BOM_xlSheet.Columns("H:H").Copy
    领料BOM_xlSheet.Columns("I:I").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    领料BOM_xlSheet.Columns("I:I").Copy
    领料BOM_xlSheet.Columns("J:J").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    领料BOM_xlSheet.Columns("K:K").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    xlApp.CutCopyMode = False
    
    '填充数据
    领料BOM_xlSheet.Cells(5, 9) = "TP1库存"
    领料BOM_xlSheet.Cells(5, 10) = "TP2库存"
    领料BOM_xlSheet.Cells(5, 11) = "TP3库存"

    领料BOM_xlSheet.Cells(5, 9).Select
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    
    '获取元件个数
    LeadPartNum = PartNum(3)
    SmtPartNum = PartNum(4)
    OtherPartNum = PartNum(5)
    
    '删除打样元件
    If DelSamplePart(领料BOM_xlSheet) = False Then
        GoTo ErrorHandle
    End If
    
    领料BOM_xlBook.Save
    
    '重新编号
    If ReNum(领料BOM_xlSheet) = False Then
        GoTo ErrorHandle
    End If
        
    领料BOM_xlBook.Save
    领料BOM_xlBook.Close (True) '关闭工作簿
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    BomAdjust = True
    Exit Function

ErrorHandle:
    
    领料BOM_xlBook.Close (True) '关闭工作簿
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    BomAdjust = False
    
End Function

Function DelSamplePart(xlSheet As Excel.Worksheet) As Boolean
    Dim rngStart        As Range
    Dim rngEND          As Range
    
    '删除打烊物料 料号存在 但是属于12345xxxxx 或xxxxx xxxxx类型
    With xlSheet.Cells
        Set rngStart = .Find("SMT元件", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngStart Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOM模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
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
             And xlSheet.Cells(i, 2) <> "SMT元件" _
             And xlSheet.Cells(i, 2) <> "DIP元件" _
             And xlSheet.Cells(i, 2) <> "其他元件" _
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
    '重新编号
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    Dim rngEND          As Range
    
    '删除打烊物料 料号存在 但是属于12345xxxxx 或xxxxx xxxxx类型
    With xlSheet.Cells
        Set rngSMT = .Find("SMT元件", lookin:=xlValues)
        Set rngLEAD = .Find("DIP元件", lookin:=xlValues)
        Set rngOther = .Find("其他元件", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOM模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
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

    Process ProcNum, "分析tsv文件信息..."
    
    On Error GoTo ErrorHandle
    
    Dim MSxlApp As Excel.Application
    Dim MSxlBook As Excel.Workbook
    Dim MSxlSheet As Excel.Worksheet
    
    Set MSxlApp = CreateObject("Excel.Application") '创建EXCEL对象
    Set MSxlBook = MSxlApp.Workbooks.Open(TmpBomFilePath) '打开已经存在的BOM模板
    MSxlApp.Visible = False  '设置EXCEL对象可见（或不可见）
    
    Set MSxlSheet = MSxlBook.Worksheets(1) '设置活动工作表
    
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
    
    '适应不同的tsv文件编码
    tsvcode = GetSetting(App.EXEName, "tsvEncoder", "tsv文件编码", "UTF-8")
    
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
    
    '转换编码格式
    If UEFSaveTextFile(tsvFilePath & "_ansi.tsv", UEFLoadTextFile(tsvFilePath, tsvDefFmt), False, UEF_ANSI) = False Then
        MsgBox "tsv文件读取转换错误！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
        GoTo ErrorHandle
    End If
        
    FileContents = UEFLoadTextFile(tsvFilePath & "_ansi.tsv", UEF_Auto)
    Kill tsvFilePath & "_ansi.tsv"       '删除中间文件
    fileinfo = Split(FileContents, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组
    
    '获取库存类型
    Dim SelStorage As String
    Dim StorageNum As Integer
    If InStr(TmpBomFilePath, "领料BOM") > 1 Then
        SelStorage = GetSetting(App.EXEName, "SelectStorage", "库存类型", "TP1")
        
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
    
    '序号    物料    状态    描述    单位    替代关系    总可 用量   近期 可用
    '0       1       2       3       4       5           6           7
    j = 0
    
    For j = 1 To UBound(fileinfo) - 1
        bomstr = Split(fileinfo(j), vbTab)
        
        Process j * 3 / UBound(fileinfo) + ProcNum + 1, "分析物料---[" & bomstr(1) & "]..."
        
        '定位各种元件位置
        With MSxlSheet.Cells
            Set rngNum = .Find(bomstr(1), lookin:=xlValues)
            If rngNum Is Nothing Then
                'MsgBox ("找不到" & bomstr(1) & "料号的对应，可能是指代")
                For m = 1 To Len(bomstr(5))
                    For n = 1 To Len(bomstr(5))
                        tPartNum = Mid(bomstr(5), m, n)
                        If IsNumeric(tPartNum) = True And Len(tPartNum) = Len(bomstr(1)) Then
                            With MSxlSheet.Cells
                                Set rngZD = .Find(tPartNum, lookin:=xlValues)
                                    If rngZD Is Nothing Then
                                        MsgBox "找不到" & bomstr(1) & "料号的对应,请更新tsv文件", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "警告"
                                        GoTo ErrorHandle
                                    Else
                                        'MSxlSheet.Cells(rngZD.Row, 4) = MSxlSheet.Cells(rngZD.Row, 4) & "单代" & bomstr(1) & vbCrLf
                                    End If
                            End With
                        End If
                    Next
                Next
            Else
                '添加物料描述段
                MSxlSheet.Cells(rngNum.Row, 3) = bomstr(3)
                
                '领料BOM中需要添加近期可用量的说明 bomstr(7) 为近期可用量
                If InStr(TmpBomFilePath, "领料BOM") > 1 Then
                    
                    If StorageNum >= 1 And StorageNum <= 3 Then
                        If bomstr(7) = "0" Then
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8).Font.Size = 8
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8).Interior.Color = 52479
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8) = "近期可用量为" & bomstr(7)
                        ElseIf InStr(bomstr(7), "-") = 1 Then
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8).Interior.Color = 52479
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8) = bomstr(7) '近期可用量为负值
                        Else
                            MSxlSheet.Cells(rngNum.Row, StorageNum + 8) = bomstr(7) '近期可用量
                        End If
                    Else
                        MsgBox "请选择库存！"
                        GoTo ErrorHandle
                    End If
                End If
            End If
        End With
    Next j
        
    MSxlBook.Close (True) '关闭工作簿
    MSxlApp.Quit '结束EXCEL对象
    Set MSxlApp = Nothing '释放xlApp对象
    
    ImportTSV = True
    Exit Function

ErrorHandle:
    MSxlBook.Close (True) '关闭工作簿
    MSxlApp.Quit '结束EXCEL对象
    Set MSxlApp = Nothing '释放xlApp对象
    
    ImportTSV = False
    
End Function


