Attribute VB_Name = "mInitBOM"
Option Explicit

Public BomItemNumber   As Integer 'Bom元素定位信息
Public BomPartNumber   As Integer
Public BomValue        As Integer
Public BomQuantity     As Integer
Public BomPartRef      As Integer
Public BomPCBfootprint As Integer

Public PartNum(6)      As Integer '所有元件数信息

Public ProjectDir      As String  '保存上次打开的目录
Public ItemName        As String  '保存上次打开的目录

Public BomFilePath     As String  '原始文件完整名
Public SaveAsPath      As String  'BOM保存的文件路径
Public tsvFilePath     As String  'tsv文件路径信息

Function BuildProjectPath(srcPath As String)
    '集中生成所有需要的目录信息，在整个工程中，仅此可写入这些路径
    Dim tmpPath As String
    BomFilePath = srcPath
    tmpPath = Right(BomFilePath, Len(BomFilePath) - InStrRev(BomFilePath, "\"))
    ProjectDir = Left(BomFilePath, InStrRev(BomFilePath, "\") - 1) & "\"
    tmpPath = ProjectDir & "BOM\" & tmpPath
    SaveAsPath = Left(tmpPath, InStrRev(tmpPath, ".") - 1)
    '在工程目录下创建BOM目录
    If Dir(ProjectDir & "BOM\") = "" Then
        MkDir ProjectDir & "BOM\"
    End If
    SaveSetting App.EXEName, "ProjectDir", "上次工作目录", ProjectDir
End Function

Function ClearPath()
    '集中生成所有需要的目录信息，在整个工程中，仅此可写入这些路径
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
    fileinfo = Split(FileContents, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组
    
    'j表示源文件BOM中的行
    'i表示行的某一列（用tab分割的）
    'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint
    '0-----------1-----------2-------3-----------4---------------5------------
    '注意orCAD导出的序列可能不与上面一致  因此需要定位
    Process 3, "读取.BOM文件信息..."
    
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
        '    MsgBox "BOM文件格式错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
        '    ReadBomFile = False
        '    Exit Function
        'End If
    Next
    
    If BomItemNumber = -1 Or BomPartNumber = -1 Or BomValue = -1 Or BomQuantity = -1 Or BomPartRef = -1 Or BomPCBfootprint = -1 Then
        MsgBox "BOM文件格式错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
        ReadBomFile = False
        Exit Function
    End If
    
    Dim IgLibInfo()        As String
    IgLibInfo = ReadLibs(LIB_NONE)
    
    Dim IsNone As Integer
    Dim j      As Integer
    For j = 1 To UBound(fileinfo) - 1
        newbomstr = Split(fileinfo(j), vbTab)
        'BOM中每个非"N"的元件必须拥有料号(可为模糊料号)
        IsNone = QueryLib(IgLibInfo, newbomstr(BomPCBfootprint))
        If IsNone = 0 Then
            If newbomstr(BomPartNumber) = "" Then
                ReadBomFile = False
                MsgBox "封装为[" & newbomstr(BomPCBfootprint) & "]料号不存在！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "BOM文件规范错误"
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
    
    '定位各种元件位置
    With xlSheet.Cells
        Set rngSMT = .Find("SMT元件", lookin:=xlValues)
        Set rngLEAD = .Find("DIP元件", lookin:=xlValues)
        Set rngOther = .Find("其他元件", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOM模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            GoTo ErrorHandle
        End If
    End With
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    '计算元件个数
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

    '定位各种元件位置
    With NBxlSheet.Cells
        Set rngNC = .Find("NC元件", lookin:=xlValues)
        Set rngDB = .Find("DBG元件", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NC元件", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Or rngEND Is Nothing Then
            MsgBox "NC_DBG模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
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

    
    '保存元件个数，供后续调用
    PartNum(0) = NcPartNum
    PartNum(1) = DbgPartNum
    PartNum(2) = DbNcPartNum
    
    PartNum(3) = LeadPartNum
    PartNum(4) = SmtPartNum
    PartNum(5) = OtherPartNum
    
    Dim msgstr As String
    msgstr = "元件信息获取成功！" & vbCrLf & vbCrLf
    msgstr = msgstr + "          贴装   元件个数为   ： " & SmtPartNum & vbCrLf
    msgstr = msgstr + "          插装   元件个数为   ： " & LeadPartNum & vbCrLf
    msgstr = msgstr + "          其他   元件个数为   ： " & OtherPartNum & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     元件个数为   ： " & NcPartNum & vbCrLf
    msgstr = msgstr + "          DBG    元件个数为   ： " & DbgPartNum & vbCrLf
    msgstr = msgstr + "          DBG_NC 元件个数为   ： " & DbNcPartNum & vbCrLf & vbCrLf
    msgstr = msgstr + "          请选择.tsv文件路径后继续操作" & vbCrLf & vbCrLf
    msgstr = msgstr + "    注意：生成的PCBA_BOM文件需要检查修改后才可供评审 "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "元件信息"
    
    xlBook.Close (True) '关闭工作簿
    NBxlBook.Close (True)
    
ErrorHandle:
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
End Function

Function BomDraft()
    On Error GoTo ErrorHandle
    
    Process 4, "创建Excel格式 BOM文档..."
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook, PLxlBook As Excel.Workbook, NBxlBook As Excel.Workbook, NonexlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet, PLxlSheet As Excel.Worksheet, NBxlSheet As Excel.Worksheet, NonexlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Process 5, "打开PCBA_BOM初稿..."
        
    'PCBA_BOM
    Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_PCBA_BOM.xls")
    Set xlSheet = xlBook.Worksheets(1)
    
    Process 6, "打开批量查询文件..."
    '打开批量资源查询xls
    Set PLxlBook = xlApp.Workbooks.Open(SaveAsPath & "_批量资源查询.xls")
    Set PLxlSheet = PLxlBook.Worksheets(1)
    
    Process 7, "打开NC_DBG元件文档..."
    '打开NC_DBG元件xls
    Set NBxlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Set NBxlSheet = NBxlBook.Worksheets(1)
    
    '打开None元件表单
    Set NonexlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
    Set NonexlSheet = NonexlBook.Worksheets(1)
    
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    
    Dim rngNC           As Range
    Dim rngDB           As Range
    Dim rngDBNC         As Range
    
    Process 8, "定位PCBA_BOM中各元件初始位置信息..."
    '定位各种元件位置
    With xlSheet.Cells
        Set rngSMT = .Find("SMT元件", lookin:=xlValues)
        Set rngLEAD = .Find("DIP元件", lookin:=xlValues)
        Set rngOther = .Find("其他元件", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Then
            MsgBox "PCBA_BOM模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            GoTo ErrorHandle
        End If
    End With
    
    Process 9, "定位NC_DBG中各元件初始位置信息..."
    '定位各种元件位置
    With NBxlSheet.Cells
        Set rngNC = .Find("NC元件", lookin:=xlValues)
        Set rngDB = .Find("DBG元件", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NC元件", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Then
            MsgBox "NC_DBG模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            GoTo ErrorHandle
        End If
    End With
    
    '========================================================
    '读取库信息
    'LEAD 库
    Process 10, "读取库文件信息..."
    
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
    fileinfo = Split(FileContents, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组
    
    Dim j               As Integer
    Dim k               As Integer
    
    On Error GoTo ErrorHandle
    For j = 1 To UBound(fileinfo) - 1
        
        bomstr = Split(fileinfo(j), vbTab)
            
         '料号后面存在两个以上换行
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
        
        Process j * 40 / UBound(fileinfo) + 10, "分析封装[" & bomstr(BomPCBfootprint) & "]..."
        
        '分析料号 要将其添加在批量查询的Excel中
        
        If IsNumeric(bomstr(BomPartNumber)) = True And bomstr(BomPartNumber) <> "" Then
            PLxlSheet.Cells(PLPartNum + 1, 1) = bomstr(BomPartNumber)
            PLPartNum = PLPartNum + 1
        End If
        
        If InStr(bomstr(BomValue), "_DBG_NC") > 0 Or bomstr(BomValue) = "DBG_NC" Then
            'DBG_NC元件
            DbNcPartNum = DbNcPartNum + 1
            xlsInsert NBxlSheet, DbNcPartNum, rngDBNC.Row, bomstr
            
        ElseIf InStr(bomstr(BomValue), "_DBG") > 0 Or bomstr(BomValue) = "DBG" Then
            'DBG元件
            DbgPartNum = DbgPartNum + 1
            xlsInsert NBxlSheet, DbgPartNum, rngDB.Row, bomstr
           
        ElseIf InStr(bomstr(BomValue), "_NC") > 0 Or bomstr(BomValue) = "NC" Then
            'NC元件
            NcPartNum = NcPartNum + 1
            xlsInsert NBxlSheet, NcPartNum, rngNC.Row, bomstr
            
        Else
        
            If bomstr(BomPCBfootprint) = "" Then
                MsgBox bomstr(BomPartNumber) & "的PCB footprint为空", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "警告"
                GoTo ErrorHandle
            End If
            
            '========================================================
            '普通元件 区分元件贴装类型
            '判断元件类型
            IsLead = QueryLib(leadLibInfo, bomstr(BomPCBfootprint))
            IsSmt = QueryLib(smtLibInfo, bomstr(BomPCBfootprint))
            IsNone = QueryLib(IgLibInfo, bomstr(BomPCBfootprint))
                
            If IsLead = 1 And IsSmt = 0 And IsNone = 0 Then
                '统计并写入LEAD元件
                LeadPartNum = LeadPartNum + 1
                xlsInsert xlSheet, LeadPartNum, rngLEAD.Row, bomstr
                
            ElseIf IsLead = 0 And IsSmt = 1 And IsNone = 0 Then
                '统计并写入SMT元件
                SmtPartNum = SmtPartNum + 1
                xlsInsert xlSheet, SmtPartNum, rngSMT.Row, bomstr
            
            ElseIf IsLead = 0 And IsSmt = 0 And IsNone = 1 Then
                '统计并写入单独的文件中 None元件
                NonePartNum = NonePartNum + 1
                '由于使用的是NC_DBG模版，因此可用rngNC.Row,而不需要重新定位
                xlsInsert NonexlSheet, NonePartNum, rngNC.Row, bomstr
                
            ElseIf IsLead = 1 And IsSmt = 1 And IsNone = 0 Then
                '特殊SMT元件 两道工序 特殊颜色标示
                SmtPartNum = SmtPartNum + 1
                xlsInsert xlSheet, SmtPartNum, rngSMT.Row, bomstr
                xlSheet.Rows(rngSMT.Row & ":" & rngSMT.Row).Interior.Color = 16737792
                
            Else
               '库文件中没有查到封装，拒绝生成BOM
                MsgBox "封装[" & bomstr(BomPCBfootprint) & "]不存在于库文件中，请更新库文件！"
                OtherPartNum = OtherPartNum + 1
                xlsInsert xlSheet, OtherPartNum, rngOther.Row, bomstr
                GoTo ErrorHandle
            End If
            
            IsLead = 0
            IsSmt = 0
            IsNone = 0

        End If
    Next j
    
    '修改机型名称
    xlSheet.Cells(2, 1) = "机型：  " & MainForm.ItemNameText.Text & "            PCBA 版本：                       半成品编号："
    If MainForm.ItemNameText.Text = "" Then
        xlSheet.Cells(2, 1).Font.ColorIndex = 5
    End If
     
    Process 50, "批量查询文件生成完毕..."
    
    Dim msgstr As String
    msgstr = ""
    msgstr = msgstr + "          可批量查询的元件个数： " & PLPartNum & vbCrLf
    msgstr = msgstr + "          贴装   元件个数为   ： " & SmtPartNum & vbCrLf
    msgstr = msgstr + "          插装   元件个数为   ： " & LeadPartNum & vbCrLf
    msgstr = msgstr + "          其他   元件个数为   ： " & OtherPartNum & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          None   元件个数为   ： " & NonePartNum & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     元件个数为   ： " & NcPartNum & vbCrLf
    msgstr = msgstr + "          DBG    元件个数为   ： " & DbgPartNum & vbCrLf
    msgstr = msgstr + "          DBG_NC 元件个数为   ： " & DbNcPartNum & vbCrLf & vbCrLf
    msgstr = msgstr + " 批量查询文件已经正确生成，请使用ERP系统查询后继续操作" & vbCrLf & vbCrLf
    msgstr = msgstr + "    注意：生成的PCBA_BOM文件需要检查修改后才可供评审 "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "元件信息"
    
    xlBook.Close (True) '关闭工作簿
    PLxlBook.Close (True)
    NBxlBook.Close (True)
    NonexlBook.Close (True)
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    '保存元件个数，供后续调用
    PartNum(0) = NcPartNum
    PartNum(1) = DbgPartNum
    PartNum(2) = DbNcPartNum
    PartNum(3) = LeadPartNum
    PartNum(4) = SmtPartNum
    PartNum(5) = OtherPartNum
    
    Process 50, "请批量查询后继续操作，选择tsv文件..."
    Exit Function
    
ErrorHandle:

    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    MsgBox "生成BOM中间文件时发生异常", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    
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
    xlBook.Close (True) '关闭工作簿
    
    '创建批量资源查询xls
    Set PLxlBook = xlApp.Workbooks.Open(App.Path & "\template\批量查询_template.xls")
    Set PLxlSheet = PLxlBook.Worksheets(1)
    PLxlBook.SaveAs (SaveAsPath & "_批量资源查询.xls")
    PLxlBook.Close (True)
    

    '创建NC_DBG元件xls
    Set NBxlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
    Set NBxlSheet = NBxlBook.Worksheets(1)
    NBxlBook.SaveAs (SaveAsPath & "_NC_DBG.xls")
    
    NBxlBook.Close (True)
    
    Set NonexlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
    NonexlBook.SaveAs (SaveAsPath & "_None_PartRef.xls")
    NonexlBook.Worksheets(1).Name = "None元件"
    
    Set NonexlSheet = NonexlBook.Worksheets(1)
    
    With NonexlSheet.Cells
        Set rngNC = .Find("NC元件", lookin:=xlValues)
        Set rngDB = .Find("DBG元件", lookin:=xlValues)
        Set rngDBNC = .Find("DBG_NC元件", lookin:=xlValues)
        If rngNC Is Nothing Or rngDB Is Nothing Then
            MsgBox "NC_DBG模板错误", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            End
        End If
    End With
    
    '调整NoneShleet
    NonexlSheet.Cells(rngNC.Row, 2) = "None"
    NonexlSheet.Rows(rngDB.Row & ":" & rngDBNC.Row + 1).Delete

    NonexlBook.Close (True)
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    Exit Function
    
ErrorHandle:
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    MsgBox "创建BOM中间文件时发生异常", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    

End Function

