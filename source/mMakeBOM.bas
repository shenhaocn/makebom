Attribute VB_Name = "mMakeBOM"
Option Explicit

'BOM类型
Public Enum BomType

BOM_ALL = 0
BOM_NCDBG
BOM_NONE

BOM_预
BOM_领料
BOM_调试
BOM_生产

End Enum

'BMF文件编码格式
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
    
    '首行不需要加入
    If ItemNum > 1 Then
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Insert
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Interior.Pattern = xlNone '去除颜色等格式 修正显示bug
    End If
    
    xlSheet.Cells(ItemNum + Row, 1) = ItemNum
    xlSheet.Cells(ItemNum + Row, 2) = insertStr(BMF_PartNum)
    xlSheet.Cells(ItemNum + Row, 3) = insertStr(BMF_Description)
    xlSheet.Cells(ItemNum + Row, 5) = insertStr(BMF_Quantity)
    xlSheet.Cells(ItemNum + Row, 6) = insertStr(BMF_PartRef)
    xlSheet.Cells(ItemNum + Row, 7) = insertStr(BMF_PcbFB)
    xlSheet.Cells(ItemNum + Row, 8) = insertStr(BMF_Value)
    
    '是否添加库存信息？
    If OrgEnable = True Then
        If insertStr(BMF_TP1) = "-" Then
            xlSheet.Cells(ItemNum + Row, 9) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 9) = insertStr(BMF_TP1)
            If insertStr(BMF_TP1) = "0" Or InStr(insertStr(BMF_TP1), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 9).Interior.Color = 52479 '以强调的颜色显示
            End If
        End If
        
        If insertStr(BMF_TP2) = "-" Then
            xlSheet.Cells(ItemNum + Row, 10) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 10) = insertStr(BMF_TP2)
            If insertStr(BMF_TP2) = "0" Or InStr(insertStr(BMF_TP2), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 10).Interior.Color = 52479 '以强调的颜色显示
            End If
        End If
            
        If insertStr(BMF_TP3) = "-" Then
            xlSheet.Cells(ItemNum + Row, 11) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 11) = insertStr(BMF_TP3)
            If insertStr(BMF_TP3) = "0" Or InStr(insertStr(BMF_TP3), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 11).Interior.Color = 52479 '以强调的颜色显示
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

'根据选项创建BOM 并调整格式 为填充数据做准备
Function ExcelCreate(bt_value As BomType)

    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    '=====================================================================================
    '预BOM：Capture中导出的BOM除None元件、NC元件、DBG元件、DBG_NC元件之外的所有元件的集合。
    '=====================================================================================
    If bt_value = BOM_预 Then
        'PCBA_BOM
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_预BOM_BMF.xls")
        
        xlBook.Close (True) '关闭工作簿
    End If
    
    '=====================================================================================
    'NC_DBG元件xls
    '=====================================================================================
    If bt_value = BOM_NCDBG Then
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_NC_DBG.xls")
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'None元件xls
    '=====================================================================================
    If bt_value = BOM_NONE Then
    
        Dim rngNC           As Range
        Dim rngDB           As Range
        Dim rngDBNC         As Range
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        xlBook.SaveAs (SaveAsPath & "_None_PartRef.xls")
        xlBook.Worksheets(1).Name = "None元件"
        
        Set xlSheet = xlBook.Worksheets(1)
        
        With xlSheet.Cells
            Set rngNC = .Find("NC元件", lookin:=xlValues)
            Set rngDB = .Find("DBG元件", lookin:=xlValues)
            Set rngDBNC = .Find("DBG_NC元件", lookin:=xlValues)
            If rngNC Is Nothing Or rngDB Is Nothing Then
                MsgBox "NC_DBG模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
                End
            End If
        End With
        
        '调整NoneShleet
        xlSheet.Cells(rngNC.Row, 2) = "None"
        xlSheet.Rows(rngDB.Row & ":" & rngDBNC.Row + 1).Delete
    
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    '领料BOM ： 预BOM + DBG元件 - 新打样物料+ 物料库存信息。
    '=====================================================================================
    If bt_value = BOM_领料 Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_领料BOM.xls")
         
        '调整领料BOM格式
        '在领料BOM中插入列(I：TP1库存) (J：TP2库存) (K：TP3库存)（需选择库存信息）
        '调整表格适应新添加的内容
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
        
        '填充数据
        xlSheet.Cells(5, 9) = "TP1库存"
        xlSheet.Cells(5, 10) = "TP2库存"
        xlSheet.Cells(5, 11) = "TP3库存"

        xlSheet.Cells(5, 9).Select
         
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    '调试BOM ： 预BOM + DBG元件 集合。
    '=====================================================================================
    If bt_value = BOM_调试 Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlSheet.SaveAs (SaveAsPath & "_调试BOM.xls")
        
        '调整打印格式
        With xlSheet.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = 80
        End With
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    '生产BOM ： 预BOM + DBG_NC元件 集合
    '=====================================================================================
    If bt_value = BOM_生产 Then

        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_生产BOM.xls")
        
        xlBook.Close (True)
    End If
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    Exit Function
    
ErrorHandle:
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    MsgBox "创建BOM中间文件时发生异常", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    

End Function

'填充相应的数据

'=====================================================================================
'1.预BOM：Capture中导出的BOM除None元件、NC元件、DBG元件、DBG_NC元件之外的所有元件的集合。
'2.NC_DBG元件xls
'3.None元件xls
'4.领料BOM ： 预BOM + DBG元件 - 新打样物料+ 物料库存信息
'5.调试BOM ： 预BOM + DBG元件 集合
'6.生产BOM ： 预BOM + DBG_NC元件 集合
'�
'元件类型 ： 普通元件 NcDbg元件 None元件 打样物料 �
'=====================================================================================
Function CreateBOM(bt_value As BomType) As Boolean
    
    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
     
    '打开相应的文件
    Select Case bt_value
    Case BOM_NCDBG:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Case BOM_NONE:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
    Case BOM_预:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_预BOM_BMF.xls")
    Case BOM_领料:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_领料BOM.xls")
    Case BOM_调试:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_调试BOM.xls")
    Case BOM_生产:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_生产BOM.xls")
    Case Else
        GoTo ErrorHandle
    End Select
     
    Set xlSheet = xlBook.Worksheets(1)
    
    '定位各种元件位置
    Dim rngPos1       As Range '在BOM文件中表示"SMT元件"位置 在DBG、None元件代表"NC元件"位置
    Dim rngPos2       As Range '在BOM文件中表示"DIP元件"位置 在DBG、None元件代表"DBG元件"位置
    Dim rngPos3       As Range '在BOM文件中表示"其他元件"位置 在DBG、None元件代表"DBG_NC元件"位置
    With xlSheet.Cells
        
        Select Case bt_value
        Case BOM_NCDBG:
            Set rngPos1 = .Find("NC元件", lookin:=xlValues)
            Set rngPos2 = .Find("DBG元件", lookin:=xlValues)
            Set rngPos3 = .Find("DBG_NC元件", lookin:=xlValues)
            
        Case BOM_NONE:
            Set rngPos1 = .Find("None", lookin:=xlValues)
            Set rngPos2 = .Find("None", lookin:=xlValues)
            Set rngPos3 = .Find("None", lookin:=xlValues)
    
        Case BOM_预, BOM_领料, BOM_调试, BOM_生产:
            Set rngPos1 = .Find("SMT元件", lookin:=xlValues)
            Set rngPos2 = .Find("DIP元件", lookin:=xlValues)
            Set rngPos3 = .Find("其他元件", lookin:=xlValues)
            
        Case Else
            GoTo ErrorHandle
        End Select
        
        If rngPos1 Is Nothing Or rngPos2 Is Nothing Or rngPos3 Is Nothing Then
            MsgBox "模板错误-元件位置定位错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            GoTo ErrorHandle
        End If
    End With
    
    
    '=======================================================
    '开始读取bmf文件并填充相应的内容
    '=======================================================
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    Dim bmfAtom()       As String
    
    Dim ItemNum1        As Integer 'NC元件 或 None元件 或 SMT元件数
    Dim ItemNum2        As Integer 'DBG元件 或 None元件 或 DIP元件数
    Dim ItemNum3        As Integer 'DBG_NC元件 或 None元件 或 其他元件数
    
    Dim i               As Integer
    Dim OrgEnable       As Boolean
    
    ItemNum1 = 0
    ItemNum2 = 0
    ItemNum3 = 0
    
    '是否添加库存信息？
    If bt_value = BOM_领料 Then
        OrgEnable = True
    Else
        OrgEnable = False
    End If
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '遍历bmf文件
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        
        If bmfAtom(BMF_PcbFB) = "" Or bmfAtom(BMF_Value) = "" Then
                MsgBox "第" & bmfAtom(BMF_ItemNum) & "项元件信息不完整", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "警告"
                GoTo ErrorHandle
        End If
        
        If InStr(bmfAtom(BMF_Value), "_DBG_NC") > 0 Or bmfAtom(BMF_Value) = "DBG_NC" Then
            If bt_value = BOM_NCDBG Then
                'DBG_NC元件
                ItemNum3 = ItemNum3 + 1
                xlsInsert xlSheet, ItemNum3, rngPos3.Row, bmfAtom, OrgEnable
            End If
            
        ElseIf InStr(bmfAtom(BMF_Value), "_DBG") > 0 Or bmfAtom(BMF_Value) = "DBG" Then
            If bt_value = BOM_NCDBG Then
                'DBG元件
                ItemNum2 = ItemNum2 + 1
                xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            End If
           
        ElseIf InStr(bmfAtom(BMF_Value), "_NC") > 0 Or bmfAtom(BMF_Value) = "NC" Then
            If bt_value = BOM_NCDBG Then
                'NC元件
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
            
            If bt_value = BOM_调试 Or BOM_领料 Or BOM_生产 Or BOM_预 Then
                '========================================================
                '普通元件 区分元件贴装类型 先不区分这四个BOM 后续调整
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
                    '库文件中没有查到封装，拒绝生成BOM
                    MsgBox "未知封装[" & bmfAtom(BMF_PcbFB) & "]，请更新库文件！"
                    GoTo ErrorHandle
                End Select
            End If
            
        End If
        
    Next i
    
    If bt_value = BOM_调试 Or BOM_领料 Or BOM_生产 Or BOM_预 Then
        '修改机型名称
        xlSheet.Cells(2, 1) = "机型：  " & MainForm.ItemNameText.Text & "            PCBA 版本：                       半成品编号："
        If MainForm.ItemNameText.Text = "" Then
            xlSheet.Cells(2, 1).Font.ColorIndex = 5
        End If
        
    End If
    
    '调整区分各种不同的BOM
    '预BOM：Capture中导出的BOM除None元件、NC元件、DBG元件、DBG_NC元件之外的所有元件的集合
    '领料BOM ： 预BOM + DBG元件 - 新打样物料+ 物料库存信息
    '调试BOM ： 预BOM + DBG元件 集合
    '生产BOM ： 预BOM + DBG_NC元件 集合
    'Select Case bt_value
    
    'End Select
     
    xlBook.Close (True) '关闭工作簿

    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    Exit Function
    
ErrorHandle:

    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    MsgBox "生成BOM中间文件时发生异常", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    
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



