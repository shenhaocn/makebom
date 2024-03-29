Attribute VB_Name = "mMakeBOM"
'*************************************************************************************
'**模 块 名：mMakeBOM
'**说    明：TP-LINK SMB Switch Product Line Hardware Group 版权所有2011 - 2012(C)
'**创 建 人：Shenhao
'**日    期：2011-10-31 23:37:45
'**修 改 人：
'**日    期：
'**描    述：Excel格式BOM生成
'**版    本：V3.6.3
'*************************************************************************************
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


'*************************************************************************
'**函 数 名：xlsInsert
'**输    入：xlSheet(Excel.Worksheet) -
'**        ：ItemNum(Integer)         -
'**        ：Row(Long)                -
'**        ：insertStr()(String)      -
'**        ：OrgEnable(Boolean)       -
'**输    出：无
'**功能描述：在对应的Sheet的对应位置上添加一行
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-01 00:37:31
'**修 改 人：
'**日    期：
'**版    本：V3.6.7
'*************************************************************************
Function xlsInsert(xlSheet As Excel.Worksheet, ItemNum As Integer, Row As Long, insertStr() As String, OrgEnable As Boolean, Optional NoteStr As String)
    
    '首行不需要加入
    If ItemNum > 1 Then
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Insert
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Interior.Pattern = xlNone '去除颜色等格式 修正显示bug
    End If
    
    xlSheet.Cells(ItemNum + Row, 1) = ItemNum
    xlSheet.Cells(ItemNum + Row, 2) = insertStr(BMF_PartNum)
    xlSheet.Cells(ItemNum + Row, 3) = insertStr(BMF_Description)
    If NoteStr <> vbNullString Then
        xlSheet.Cells(ItemNum + Row, 4) = NoteStr
        xlSheet.Cells(ItemNum + Row, 4).Font.ColorIndex = 5
    End If
    xlSheet.Cells(ItemNum + Row, 5) = insertStr(BMF_Quantity)
    xlSheet.Cells(ItemNum + Row, 6) = insertStr(BMF_PartRef)
    xlSheet.Cells(ItemNum + Row, 7) = insertStr(BMF_PcbFB)
    xlSheet.Cells(ItemNum + Row, 8) = insertStr(BMF_Value)
    
    '是否添加库存信息？
    If OrgEnable = True Then
    
        '添加TP1库存信息
        If insertStr(BMF_TP1) = "-" Then
            xlSheet.Cells(ItemNum + Row, 9) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 9) = insertStr(BMF_TP1)
            If insertStr(BMF_TP1) = "0" Or InStr(insertStr(BMF_TP1), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 9).Interior.Color = 52479 '以强调的颜色显示
            End If
        End If
        
        '添加TP2库存信息
        If insertStr(BMF_TP2) = "-" Then
            xlSheet.Cells(ItemNum + Row, 10) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 10) = insertStr(BMF_TP2)
            If insertStr(BMF_TP2) = "0" Or InStr(insertStr(BMF_TP2), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 10).Interior.Color = 52479 '以强调的颜色显示
            End If
        End If
            
        '添加TP3库存信息
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

'排序数组（可以是字符串数组）：BigToSmall=True 从大到小，否则 从小大到
Function ZuSorted(Zu() As Variant, prefixStr As String, Optional BigToSmall As Boolean) As String
   
   Dim i As Long, j As Long, s As Variant
   Dim TF As Boolean, nL As Long, nU As Long
   
   nL = LBound(Zu): nU = UBound(Zu)
   For i = nL To nU
      For j = nL To nU '       从大到小                从小大到
        If BigToSmall Then TF = Zu(j) < Zu(i) Else TF = Zu(j) > Zu(i)
        If TF Then s = Zu(i): Zu(i) = Zu(j): Zu(j) = s
      Next j
   Next i
   
   For i = nL To nU - 1
      ZuSorted = ZuSorted + prefixStr + CStr(Zu(i)) + Space(1)
   Next i
   
   ZuSorted = ZuSorted + prefixStr + CStr(Zu(nU))
   
End Function

'由于位号长度不一，常规排序方法不可行 因此需要特殊排序方法
Function RealSorted(ByRef RefStr As String, Optional BigToSmall As Boolean) As Boolean
    Dim srcStr()    As String
    Dim intSorted() As Variant
    Dim i           As Long
    Dim Index       As Long '位号后数字开始的位置
    
    Dim prefixStr   As String
    
    RealSorted = False
    
    srcStr = Split(RefStr, Space(1))
    ReDim intSorted(UBound(srcStr))
    
    Index = 0
    For i = 0 To Len(srcStr(0))
        If IsNumeric(Right(srcStr(0), Len(srcStr(0)) - i)) = True Then
            prefixStr = Left(srcStr(0), i)
            Index = i
            Exit For
        End If
    Next
    
    For i = LBound(intSorted) To UBound(intSorted)
        If IsNumeric(Right(srcStr(i), Len(srcStr(i)) - Index)) = True Then
            intSorted(i) = Val(Right(srcStr(i), Len(srcStr(i)) - Index))
        Else
            RealSorted = False
        End If
    Next i
    
    RefStr = ZuSorted(intSorted, prefixStr, BigToSmall)
    RealSorted = True
    
End Function

'添加DBG NC元件到相应的Excel BOM中 对合并的位号进行排序
Function addDbgNcPart(xlSheet As Excel.Worksheet, bmfAtom() As String, _
                      ByRef ItemNum1 As Integer, ByRef ItemNum2 As Integer, _
                      rngPos1 As Range, rngPos2 As Range, OrgEnable As Boolean)

    Dim rngNum       As Range
    Dim partRefStr() As String
    Dim tmpRefStr    As String
    Dim i            As Integer
    
    '是否需要合并元件
    With xlSheet.Cells
        Set rngNum = .Find(bmfAtom(BMF_PartNum), lookin:=xlValues)
        If rngNum Is Nothing Then
            '不需要合并 直接添加在后面
            Select Case bmfAtom(BMF_MountType)
                Case "S"
                ItemNum1 = ItemNum1 + 1
                xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Font.ColorIndex = 5

                Case "S+"
                ItemNum1 = ItemNum1 + 1
                xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable, "位号对应多个物料"
                xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Font.ColorIndex = 5
                
                Case "L"
                ItemNum2 = ItemNum2 + 1
                xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
                xlSheet.Rows((rngPos2.Row + ItemNum2) & ":" & (rngPos2.Row + ItemNum2)).Font.ColorIndex = 5

                Case "N"
                'Do Nothing

                Case Else
                '库文件中没有查到封装，拒绝生成BOM
                MsgBox "未知封装[" & bmfAtom(BMF_PcbFB) & "]，请更新库文件！"
                
            End Select
            
        Else '需要合并 修改数量(第5列) 和位号(第6列) 位号需要重新排序
            'xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            xlSheet.Cells(rngNum.Row, 5) = CInt(xlSheet.Cells(rngNum.Row, 5)) + CInt(bmfAtom(BMF_Quantity))
            xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
            
            '位号需要排序！
            tmpRefStr = xlSheet.Cells(rngNum.Row, 6) + " " + bmfAtom(BMF_PartRef)
            If RealSorted(tmpRefStr, False) = False Then
                xlSheet.Cells(rngNum.Row, 4) = "位号排序异常！"
                xlSheet.Cells(rngNum.Row, 4).Font.ColorIndex = 5
            End If
            
            xlSheet.Cells(rngNum.Row, 6) = tmpRefStr
            xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
            
        End If
    End With
End Function

'行拷贝
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
        
        '填充数据 第5行才是真正的列头描述行
        xlSheet.Cells(5, 9) = "TP1库存"
        xlSheet.Cells(5, 10) = "TP2库存"
        xlSheet.Cells(5, 11) = "TP3库存"

        xlSheet.Cells(5, 1).Select
         
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

    xlBook.Close (True) '关闭工作簿
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
        Case BOM_NCDBG

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
        Case BOM_NONE

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
        Case BOM_预

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_预BOM_BMF.xls")
        Case BOM_领料

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_领料BOM.xls")
        Case BOM_调试

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_调试BOM.xls")
        Case BOM_生产

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
            Case BOM_NCDBG

            Set rngPos1 = .Find("NC元件", lookin:=xlValues)
            Set rngPos2 = .Find("DBG元件", lookin:=xlValues)
            Set rngPos3 = .Find("DBG_NC元件", lookin:=xlValues)

            Case BOM_NONE

            Set rngPos1 = .Find("None", lookin:=xlValues)
            Set rngPos2 = .Find("None", lookin:=xlValues)
            Set rngPos3 = .Find("None", lookin:=xlValues)

            Case BOM_预, BOM_领料, BOM_调试, BOM_生产

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
    '需要检测tsv文件创建时间 如果时间不在3天内 库存信息添加了也是没有用的
    '已经本Function调用前检测了tsv文件时间
    If bt_value = BOM_领料 Then
        OrgEnable = True
    Else
        OrgEnable = False
    End If

    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If

    '读取bmf文件 并按行分割为数组
    bmfBomLine = Split(GetFileContents(BmfFilePath), vbCrLf)

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

            If bt_value = BOM_调试 Or bt_value = BOM_领料 Or _
               bt_value = BOM_生产 Or bt_value = BOM_预 Then
                '========================================================
                '普通元件 区分元件贴装类型 先不区分这四个BOM 后续调整
                Select Case bmfAtom(BMF_MountType)
                    Case "S"
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable

                    Case "S+"
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable, "位号对应多个物料"

                    Case "L"
                    ItemNum2 = ItemNum2 + 1
                    xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable

                    Case "N"
                    'Do Nothing

                    Case Else
                    '库文件中没有查到封装，拒绝生成BOM
                    MsgBox "未知封装[" & bmfAtom(BMF_PcbFB) & "]，请更新库文件！"
                    GoTo ErrorHandle
                    
                End Select
            End If

        End If

    Next i

    If bt_value = BOM_调试 Or bt_value = BOM_领料 Or _
       bt_value = BOM_生产 Or bt_value = BOM_预 Then
        '修改BOM标题
        Select Case bt_value
        Case BOM_调试
            xlSheet.Cells(1, 1) = "调试BOM初稿"
        Case BOM_领料
            xlSheet.Cells(1, 1) = "领料BOM初稿"
        Case BOM_生产
            xlSheet.Cells(1, 1) = "生产BOM初稿"
        Case BOM_预
            xlSheet.Cells(1, 1) = "预BOM初稿"
        Case Else
            xlSheet.Cells(1, 1) = "预BOM初稿"
        End Select
        
        '修改机型名称
        xlSheet.Cells(2, 1) = "机型：  " & MainForm.ItemNameText.Text & "            PCBA 版本：                       半成品编号："
        If MainForm.ItemNameText.Text = "" Then
            xlSheet.Cells(2, 1).Font.ColorIndex = 5
        End If

    End If

    If bt_value = BOM_调试 Or bt_value = BOM_领料 Or bt_value = BOM_生产 Then
        '调整区分各种不同的BOM
        '预BOM：Capture中导出的BOM除None元件、NC元件、DBG元件、DBG_NC元件之外的所有元件的集合
        '领料BOM ： 预BOM + DBG元件 - 新打样物料 + 物料库存信息
        '调试BOM ： 预BOM + DBG元件 集合
        '生产BOM ： 预BOM + DBG_NC元件 集合
        
        '重新遍历bmf文件 根据调整相应BOM信息
         For i = 1 To UBound(bmfBomLine) - 1
            bmfAtom = Split(bmfBomLine(i), vbTab)
    
            If InStr(bmfAtom(BMF_Value), "_DBG_NC") > 0 Or bmfAtom(BMF_Value) = "DBG_NC" Then
            
                If bt_value = BOM_生产 Then
                    addDbgNcPart xlSheet, bmfAtom, ItemNum1, ItemNum2, rngPos1, rngPos2, OrgEnable
                End If
    
            ElseIf InStr(bmfAtom(BMF_Value), "_DBG") > 0 Or bmfAtom(BMF_Value) = "DBG" Then
            
                If bt_value = BOM_调试 Or bt_value = BOM_领料 Then
                    addDbgNcPart xlSheet, bmfAtom, ItemNum1, ItemNum2, rngPos1, rngPos2, OrgEnable
                End If
    
            End If
    
        Next i
        
        '删除打样元件
        If bt_value = BOM_领料 Then
            If DelSamplePart(xlSheet) = False Then
                GoTo ErrorHandle
            End If
        End If
    
        '重新编号
        If ReNum(xlSheet) = False Then
            GoTo ErrorHandle
        End If
    End If
    
    '是否经过BomChecker？
    '经过模版的Excel文件格式应该是OK的
    
    '结束所有调整和创建 添加
    xlBook.Close (True) '关闭工作簿

    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象

    Exit Function

ErrorHandle:
    
    xlBook.Close (True) '关闭工作簿
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象

    MsgBox "生成BOM中间文件时发生异常", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"

End Function

'删除打样物料
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
    Dim PartNum    As Integer   '伪元件个数
    Dim DelRows()  As Integer   '记录要删除的行号
    
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

'重新编号
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

'检查BOM
'a.调整位号对应的列宽?行高使之全部显示
'b.元件数和位号数一致
'c.添加备注系列
'c.1.Flash增加备注: 需预编程
'c.2.烧录软件添加备注: SMT用
'c.3.工具软件添加备注："测试阶段用"
Function BomChecker(ExcelBomFilePath As String)
    On Error GoTo ErrorHandle

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    '定位元素位置 定位列的位置
    Dim rngNum    As Range '在BOM文件中表示"序号"的位置
    Dim rngDcp    As Range '在BOM文件中表示"规格型号"的位置
    Dim rngNote   As Range '在BOM文件中表示"辅助说明"的位置
    Dim rngQty    As Range '在BOM文件中表示"数量"的位置
    Dim rngRef    As Range '在BOM文件中表示"位号"的位置
    
    Dim usedRow   As Integer  '总行数
    Dim usedCol   As Integer  '总列数
    
    Dim BomAtom() As String
    Dim tmpRefStr As String
    
    Dim AutoFitFlag As Boolean
    
    If CInt(GetExcelVer) > 14 Then
        AutoFitFlag = True
    Else
        AutoFitFlag = False
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Process 10, "打开文件...."
    
    '打开文件
    Set xlBook = xlApp.Workbooks.Open(ExcelBomFilePath)
    '先检查第第一个WorkSheet
    Set xlSheet = xlBook.Worksheets(1)
    
    Process 15, "验证文件...."
    With xlSheet.Cells

        Set rngNum = .Find("序号", lookin:=xlValues)
        Set rngDcp = .Find("规格型号", lookin:=xlValues)
        Set rngNote = .Find("辅助说明", lookin:=xlValues)
        Set rngQty = .Find("数量", lookin:=xlValues)
        Set rngRef = .Find("位号", lookin:=xlValues)
        
    End With
    
    If rngNum Is Nothing Or rngNote Is Nothing Or rngQty Is Nothing Or rngRef Is Nothing Then
        '检查第第二个WorkSheet
        Set xlSheet = xlBook.Worksheets(2)
        With xlSheet.Cells
    
            Set rngNum = .Find("序号", lookin:=xlValues)
            Set rngDcp = .Find("规格型号", lookin:=xlValues)
            Set rngNote = .Find("辅助说明", lookin:=xlValues)
            Set rngQty = .Find("数量", lookin:=xlValues)
            Set rngRef = .Find("位号", lookin:=xlValues)
        End With
        
        If rngNum Is Nothing Or rngNote Is Nothing Or rngQty Is Nothing Or rngRef Is Nothing Then
                '仅支持前两个Sheet 其他的不支持
                MsgBox "BOM元素位置定位错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
                GoTo ErrorHandle
            End If
    End If
    
    Process 20, "元素位置完成...."
    
    usedRow = xlSheet.UsedRange.Rows.Count
    usedCol = xlSheet.UsedRange.Columns.Count
    '===============================================================================================
    '开始Check
    '===============================================================================================
    
    Dim intFontSize   As Integer  '字体大小  单位：磅
    Dim intRowNum     As Integer  '单词宽度  单位：字符
    Dim calRowHeight  As Integer  '计算所得行高
    Dim calColWidth   As Integer  '计算所得列宽
    
    Dim j As Integer
    For j = rngNum.Row + 1 To usedRow
    
    'a.调整位号对应的列宽?行高使之全部显示
        '先整理内容 去除软回车 多余的空格等等
        xlSheet.Cells(j, rngRef.Column) = clearRefStr(xlSheet.Cells(j, rngRef.Column))
        
        '调整位号排序
        Process j * 70 / (usedRow - rngNum.Row) + 20, "排序" & "[" & j & "]" & "行位号..."
        If xlSheet.Cells(j, rngRef.Column) <> "" Then
            tmpRefStr = xlSheet.Cells(j, rngRef.Column)
            If RealSorted(tmpRefStr, False) = True Then
                xlSheet.Cells(j, rngRef.Column) = tmpRefStr
            Else
                xlSheet.Cells(j, rngNote.Column) = "位号重排序异常！"
                xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
            End If
        End If
        
        '调整位号对应的列宽?行高使之全部显示
        Process j * 70 / (usedRow - rngNum.Row) + 20, "调整" & "[" & j & "]" & "行位号..."
        If xlSheet.Cells(j, rngRef.Column) <> "" Then
            '精确获取字体大小 + 段落高度
            intFontSize = LineHeight(xlSheet.Cells(j, rngRef.Column).Font.Size)
            intRowNum = calRowNum(xlSheet.Cells(j, rngRef.Column), xlSheet.Cells(j, rngRef.Column).ColumnWidth)
            '计算所需的合适的行高 列宽
            'Debug.Print "字体大小为：" & intFontSize & "    单词宽度为：" & intWordWidth
            If xlSheet.Cells(j, rngRef.Column).Height / intFontSize < intRowNum Then
                '需要计算手动调整
                AutoFitFlag = False
                
                '首先尝试按目前的列宽调整行高
                calColWidth = xlSheet.Cells(j, rngRef.Column).ColumnWidth
                calRowHeight = intRowNum * intFontSize
                
                Do While calRowHeight > 408
                    calColWidth = calColWidth + 1
                    intRowNum = calRowNum(xlSheet.Cells(j, rngRef.Column), calColWidth)
                    calRowHeight = intRowNum * intFontSize
                Loop
            Else
                AutoFitFlag = True
            End If
            
            If AutoFitFlag Then
                '可自动适应 OK
                With xlSheet.Cells(j, rngRef.Column)
                    .WrapText = True   '自动换行
                    .Rows.AutoFit      '自适应行高
                End With
                '自此大多数显示应该都正确了
                If xlSheet.Cells(j, rngRef.Column).Height > 408 Then
                    xlSheet.Cells(j, rngRef.Column).Columns.AutoFit
                    xlSheet.Cells.EntireRow.AutoFit
                End If
            Else
                '需要手动计算 手动调整
                xlSheet.Cells(j, rngRef.Column).RowHeight = calRowHeight
                If calColWidth <> xlSheet.Cells(j, rngRef.Column).ColumnWidth Then
                    '行高需要调整 同时需要将其他的行自动适应行高
                    xlSheet.Cells(j, rngRef.Column).ColumnWidth = calColWidth
                    xlSheet.Cells.EntireRow.AutoFit
                End If
                '颜色标记经过调整的单元格
                xlSheet.Cells(j, rngRef.Column).Font.ColorIndex = 3 '标记为红色 (5为蓝色)
            End If
        End If
        
        
    'b.元件数和位号数一致
        Process j * 70 / (usedRow - rngNum.Row) + 20, "检查" & "[" & j & "]" & "行元件数..."
        If xlSheet.Cells(j, rngQty.Column) <> "" And xlSheet.Cells(j, rngRef.Column) <> "" Then
           If InStr(xlSheet.Cells(j, rngRef.Column), "优先级") = 0 Then
               BomAtom = Split(xlSheet.Cells(j, rngRef.Column), Space(1))
               If CInt(xlSheet.Cells(j, rngQty.Column)) <> (UBound(BomAtom) + 1) Then
                  '不相等颜色标记
                  xlSheet.Cells(j, rngQty.Column).Interior.Color = 255 '以强调的颜色显示
                  xlSheet.Cells(j, rngRef.Column).Interior.Color = 255 '以强调的颜色显示
                  xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + _
                                                     "元件数[" & CInt(xlSheet.Cells(j, rngQty.Column)) & "]" & _
                                                     "位号数[" & (UBound(BomAtom) + 1) & "]不相等！"
                  xlSheet.Cells(j, rngNote.Column).Interior.Color = 255 '以强调的颜色显示
                  xlSheet.Cells(j, rngNote.Column).Font.Size = 8
               End If
           End If
        End If
        
        
    'c.添加备注系列
        Process j * 70 / (usedRow - rngNum.Row) + 20, "检查" & "[" & j & "]" & "行辅助说明..."
        'c.1.Flash增加备注: 需预编程
        If InStr(xlSheet.Cells(j, rngDcp.Column), "FLASH") > 0 And _
           InStr(xlSheet.Cells(j, rngNote.Column), "预编程") = 0 Then
            xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "需预编程"
            xlSheet.Cells(j, rngNote.Column).Font.Size = 10
            xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
        End If
        'c.2.烧录软件添加备注: SMT用
        If InStr(xlSheet.Cells(j, rngDcp.Column), "烧录软件") > 0 And _
           InStr(xlSheet.Cells(j, rngNote.Column), "用") = 0 Then
            xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "SMT用"
            xlSheet.Cells(j, rngNote.Column).Font.Size = 10
            xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
        End If
        'c.3.工具软件添加备注："测试阶段用"
        If InStr(xlSheet.Cells(j, rngDcp.Column), "工具软件") > 0 And _
           InStr(xlSheet.Cells(j, rngNote.Column), "用") = 0 Then
            xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "测试阶段用"
            xlSheet.Cells(j, rngNote.Column).Font.Size = 10
            xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
        End If
    Next j
        
    Process 95, "所有检测结束！"
    '结束所有调整和创建 添加
    xlBook.Close (True) '关闭工作簿

    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    
    Process 100, "检测完成！"
    MsgBox "BOM格式整理完毕！" & vbCrLf & vbCrLf & "BOM检查结束！", vbMsgBoxSetForeground + vbOKOnly + vbInformation, "提示"
    
    Exit Function
    
ErrorHandle:
    
    xlBook.Close (True) '关闭工作簿
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象

    MsgBox "打开Excel格式BOM文件时发生异常！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
End Function

'精确计算所需的行数(根据列宽计算，列宽单位为字符)
Function calRowNum(RefStr As String, ColumnWidth As Integer) As Integer
    Dim Index As Long
    
    Do
        Index = InStrRev(RefStr, Space(1), ColumnWidth)
        RefStr = Right(RefStr, Len(RefStr) - Index)
        calRowNum = calRowNum + 1
    Loop While Index <> 0
    
End Function

'根据字体大小返回对应的因子 即返回（字体高度+行距）
Function LineHeight(FontSize As Integer) As Double
    
    Dim LineFactor As Double
    Select Case FontSize
    Case 8
        LineFactor = 1.35
    Case 9
        LineFactor = 1.28
    Case 10
        LineFactor = 1.22
    Case 11
        LineFactor = 1.25
    Case 12
        LineFactor = 1.2
    Case 14
        LineFactor = 1.36
    Case Else '不支持的字体大小
        LineFactor = 1
    End Select
    
    LineHeight = LineFactor * FontSize
    
End Function

Function clearRefStr(subRefStr As String) As String
    If subRefStr = "" Then
        Exit Function
    End If
    
    clearRefStrSub subRefStr, vbCrLf
    clearRefStrSub subRefStr, vbCr
    clearRefStrSub subRefStr, vbLf
    
    clearRefStr = subRefStr
    
End Function

Function clearRefStrSub(ByRef tmpRefStr As String, spChar As String)
    
    'vbCrLf -> Space(1)
    Do While InStr(tmpRefStr, spChar) > 0
        tmpRefStr = Replace(tmpRefStr, spChar, Space(1))
    Loop
    'Space(2)->Space(1)
    Do While InStr(tmpRefStr, Space(2))
        tmpRefStr = Replace(tmpRefStr, Space(2), Space(1))
    Loop
    '位于开始位置的Space(1)
    If InStr(tmpRefStr, Space(1)) = 1 Then
        tmpRefStr = Replace(tmpRefStr, Space(1), "", 1, 1)
    End If
    '位于结束位置的Space(1)
    Do While Right(tmpRefStr, 1) = Space(1)
        tmpRefStr = Left(tmpRefStr, Len(tmpRefStr) - 1)
    Loop
    
End Function
