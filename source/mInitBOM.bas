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
Public BmfFilePath     As String  '原始文件完整名
Public SaveAsPath      As String  'BOM保存的文件路径
Public tsvFilePath     As String  'tsv文件路径信息

'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
'0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--

'BMF文件信息编码格式
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
    '集中生成所有需要的目录信息，在整个工程中，仅此可写入这些路径
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
    
    '在工程目录下创建BOM目录
    If Dir(ProjectDir & "BOM\", vbDirectory) = "" Then
        MkDir ProjectDir & "BOM\"
    End If
    
    SaveSetting App.EXEName, "ProjectDir", "上次工作目录", ProjectDir
End Function

Function ClearPath()
    '集中生成所有需要的目录信息，在整个工程中，仅此可写入这些路径
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
    
    '创建批量查询Excel文件
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\批量查询_template.xls")
    Set xlSheet = xlBook.Worksheets(1)
    xlBook.SaveAs (SaveAsPath & "_批量资源查询.xls")
    xlBook.Close (True)
    
    '打开批量资源查询xls
    Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_批量资源查询.xls")
    Set xlSheet = xlBook.Worksheets(1)
    
    Bom = GetFileContents(BomFilePath)
    
    BomLine = Split(Bom, vbCrLf)
    
    Atom = Split(BomLine(0), vbTab)
    
    
    '将Bom的信息导入
    For i = 1 To UBound(BomLine) - 1
        Atom = Split(BomLine(i), vbTab)
        
        '分析料号 要将其添加在批量查询的Excel中
        If IsNumeric(Atom(BomPartNumber)) = True And Atom(BomPartNumber) <> "" Then
            xlSheet.Cells(PartNum + 1, 1) = Atom(BomPartNumber)
            PartNum = PartNum + 1
        End If
    Next i
    
    xlBook.Close (True) '关闭工作簿
    
    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    Process 20, "批量查询文件生成完毕..."
    
    Exit Function
    
ErrorHandle:

    xlApp.Quit '结束EXCEL对象
    Set xlApp = Nothing '释放xlApp对象
    
    MsgBox "生成BOM中间文件时发生异常", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    
End Function

Function ReadBomFile() As Boolean
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim newbomstr()     As String
    
    Dim i               As Integer
    
    FileContents = GetFileContents(BomFilePath)
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

'遍历旧Bom 检查BMF文件格式是否正确
Function CheckBmf() As Boolean
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    Dim i               As Integer
    
    '初始化返回值
    CheckBmf = False
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '遍历旧Bom 检查BMF文件格式是否正确
    For i = 0 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        If UBound(bmfAtom) = BMF_TOTAL Then
            CheckBmf = True
        End If
    Next i
    
End Function

'根据给定的字符串查找给定的列，返回给定列号的字符串
Function LookupBmfAtom(checkStr As String, checkCol As Integer, returnCol As Integer) As String
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    Dim i               As Integer
    
    '初始化返回值
    LookupBmfAtom = "-"
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '遍历旧Bom 查找对应的字符串
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        If checkStr = bmfAtom(checkCol) Then
            LookupBmfAtom = bmfAtom(returnCol)
        End If
    Next i
    
End Function

'根据给定的字符串查找给定的列，返回查找到的第一个行号
Function LookupBmfRow(checkStr As String, checkCol As Integer) As Integer
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    Dim i               As Integer
    
    '初始化返回值
    LookupBmfRow = -1
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    '遍历旧Bom 查找对应的字符串
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)
        If checkStr = bmfAtom(checkCol) Then
            LookupBmfRow = i
            Exit For
        End If
    Next i
    
End Function

'根据给定行号，给定列号，返回字符串
Function GetBmfAtom(Row As Integer, Col As Integer) As String
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String

    
    '初始化返回值
    GetBmfAtom = ""
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)
    
    bmfAtom = Split(bmfBomLine(Row), vbTab)
    GetBmfAtom = bmfAtom(Col)
    
End Function

'根据给定行号，返回一行数据
Function GetBmfLine(Row As Integer) As String
        
    Dim bmfBom          As String
    Dim bmfBomLine()    As String
    
    Dim bmfAtom()       As String
    
    '初始化返回值
    GetBmfLine = ""
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

    GetBmfLine = bmfBomLine(Row)
    
End Function

'修改给定行号，给定列号的对应的数据
Function SetBmfAtom(Row As Integer, Col As Integer, addStr As String)
             
    Dim oldBom          As String
    Dim newBomLine()    As String
    
    Dim BomAtom()       As String
    
    Dim i               As Integer
    
    oldBom = GetFileContents(BmfFilePath)
    
    newBomLine = Split(oldBom, vbCrLf)
    
    '添加列头信息
    'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
    '0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--
    BomAtom = Split(newBomLine(Row), vbTab)
    BomAtom(Col) = addStr
    
    newBomLine(Row) = ""
    
    '将旧Bom的信息导入到新Bom 遍历旧Bom
    For i = 0 To UBound(BomAtom) - 1
        newBomLine(Row) = newBomLine(Row) + BomAtom(i) + vbTab
    Next i
    
    '最后一列没有vbTab
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
    

    '读取库信息
    Process 10, "读取库文件信息..."
    
    Dim leadLibInfo()      As String
    Dim smtLibInfo()       As String
    Dim IgLibInfo()        As String
    
    leadLibInfo = ReadLibs(LIB_LEAD)
    smtLibInfo = ReadLibs(LIB_SMD)
    IgLibInfo = ReadLibs(LIB_NONE)
        
    '创建BOM中间文件
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
    
    'BMF文件是否存在 是否可以利用
    BmfExistFlag = CheckBmf
    
    oldBom = GetFileContents(BomFilePath)
    
    oldBomLine = Split(oldBom, vbCrLf)
    newBomLine = Split(oldBom, vbCrLf)
    
    For i = 0 To UBound(oldBomLine)
        newBomLine(i) = ""
    Next i
    
    '添加列头信息
    'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
    '0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--
    oldAtom = Split(oldBomLine(0), vbTab)
    newBomLine(0) = oldAtom(BomItemNumber) + vbTab + oldAtom(BomPartNumber) + vbTab
    newBomLine(0) = newBomLine(0) + oldAtom(BomValue) + vbTab + oldAtom(BomQuantity) + vbTab
    newBomLine(0) = newBomLine(0) + oldAtom(BomPartRef) + vbTab + oldAtom(BomPCBfootprint) + vbTab
    '贴装方式信息及物料描述信息段
    newBomLine(0) = newBomLine(0) + "Mount Type" + vbTab + "Description" + vbTab
    '库存信息
    newBomLine(0) = newBomLine(0) + "TP1" + vbTab + "TP2" + vbTab + "TP3"
    
    
    '将旧Bom的信息导入到新Bom 遍历旧Bom
    'On Error GoTo ErrorHandle
    For i = 1 To UBound(oldBomLine) - 1
        oldAtom = Split(oldBomLine(i), vbTab)
        
         '料号后面存在两个以上换行
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
        
        Process i * 40 / UBound(oldBomLine) + 10, "分析封装[" & oldAtom(BomPCBfootprint) & "]..."
        
        '分析料号
        If IsNumeric(oldAtom(BomPartNumber)) = True And oldAtom(BomPartNumber) <> "" Then
            PLPartNum = PLPartNum + 1
        End If
        
        '统计元件类型个数
        If InStr(oldAtom(BomValue), "_DBG_NC") > 0 Or oldAtom(BomValue) = "DBG_NC" Then
            'DBG_NC元件
            DbNcPartNum = DbNcPartNum + 1
            
        ElseIf InStr(oldAtom(BomValue), "_DBG") > 0 Or oldAtom(BomValue) = "DBG" Then
            'DBG元件
            DbgPartNum = DbgPartNum + 1
           
        ElseIf InStr(oldAtom(BomValue), "_NC") > 0 Or oldAtom(BomValue) = "NC" Then
            'NC元件
            NcPartNum = NcPartNum + 1
            
        End If
        
        
        If oldAtom(BomPCBfootprint) = "" Then
            MsgBox oldAtom(BomPartNumber) & "的PCB footprint为空", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "警告"
            Exit Function
        End If
        
        '========================================================
        '区分元件贴装类型  填充Mount Type段
        
        '判断元件类型
        IsLead = QueryLib(leadLibInfo, oldAtom(BomPCBfootprint))
        IsSmt = QueryLib(smtLibInfo, oldAtom(BomPCBfootprint))
        IsNone = QueryLib(IgLibInfo, oldAtom(BomPCBfootprint))
            
        If IsLead = 1 And IsSmt = 0 And IsNone = 0 Then
            '统计并写入LEAD元件
            LeadPartNum = LeadPartNum + 1
            newBomLine(i) = newBomLine(i) + "L" + vbTab
            
        ElseIf IsLead = 0 And IsSmt = 1 And IsNone = 0 Then
            '统计并写入SMT元件
            SmtPartNum = SmtPartNum + 1
            newBomLine(i) = newBomLine(i) + "S" + vbTab
        
        ElseIf IsLead = 0 And IsSmt = 0 And IsNone = 1 Then
            '统计并写入单独的文件中 None元件
            NonePartNum = NonePartNum + 1
            newBomLine(i) = newBomLine(i) + "N" + vbTab
            
        ElseIf IsLead = 1 And IsSmt = 1 And IsNone = 0 Then
            '特殊SMT元件 两道工序 特殊颜色标示
            SmtPartNum = SmtPartNum + 1
            newBomLine(i) = newBomLine(i) + "S+" + vbTab
            
        Else
           '库文件中没有查到封装，拒绝生成BOM
            MsgBox "封装[" & oldAtom(BomPCBfootprint) & "]不存在于库文件中，请更新库文件！"
            Exit Function
            
        End If
        
        IsLead = 0
        IsSmt = 0
        IsNone = 0
        
        '添加Description TPn等信息 如果存在旧的.bmf文件，可以从其导入，节约效率
        '贴装方式信息及物料描述信息段
        If BmfExistFlag Then
            '存在旧文件导入相关信息 仅对比料号 料号不相等 无法导入
            If IsNumeric(oldAtom(BomPartNumber)) = True And oldAtom(BomPartNumber) <> "" Then
                'BomPartNumber是keyed
                '描述信息
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_Description) + vbTab
                '库存信息
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_TP1) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_TP2) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomPartNumber), BMF_PartNum, BMF_TP3)
            
            ElseIf oldAtom(BomValue) <> "" Then
                'BomValue也是Keyed
                '描述信息
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_Description) + vbTab
                '库存信息
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_TP1) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_TP2) + vbTab
                newBomLine(i) = newBomLine(i) + LookupBmfAtom(oldAtom(BomValue), BMF_Value, BMF_TP3)
            Else
                '仅填充相应的项
                newBomLine(i) = newBomLine(i) + "-" + vbTab
                newBomLine(i) = newBomLine(i) + "-" + vbTab + "-" + vbTab + "-"
            End If
        Else
            '仅填充相应的项
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
    
    '保存元件个数，供后续调用
    PartNum(0) = NcPartNum
    PartNum(1) = DbgPartNum
    PartNum(2) = DbNcPartNum
    PartNum(3) = LeadPartNum
    PartNum(4) = SmtPartNum
    PartNum(5) = 0
    
    Process 50, "BOM中间文件生成成功..."
    
End Function

Function ImportTSV() As Boolean

    Process 51, "分析tsv文件信息..."
    
    Dim FileContents    As String
    Dim fileinfo()      As String
    Dim bomstr()        As String
    
    Dim i               As Integer
    Dim FindRow         As Integer
    
    '自动适应不同的tsv文件编码
    FileContents = UEFLoadTextFile(tsvFilePath, UEF_AUTO)
    fileinfo = Split(FileContents, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组
    
    '获取库存类型
    Dim SelStorage As String
    
    SelStorage = GetSetting(App.EXEName, "SelectStorage", "库存类型", "TP1")
        
    
    '序号    物料(编码)    状态    描述    单位    替代关系    总可 用量   近期 可用
    '0       1             2       3       4       5           6           7
    
    For i = 1 To UBound(fileinfo) - 1
        bomstr = Split(fileinfo(i), vbTab)
        
        Process i * 3 / UBound(fileinfo) + 51 + 1, "分析物料  [" & bomstr(1) & "]..."
        
        '查找并填入相关信息
        FindRow = LookupBmfRow(bomstr(1), BMF_PartNum)
        If FindRow > 0 Then
            '物料描述
            SetBmfAtom FindRow, BMF_Description, bomstr(3)
            
            '库存信息
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
    
    Process 60, "BOM中间文件生成完毕..."
 
End Function

Function BmfToAnsi() As Boolean
    
    '转换编码格式
    If UEFSaveTextFile(BmfFilePath, UEFLoadTextFile(BmfFilePath, UEF_AUTO), False, UEF_ANSI) = False Then
        MsgBox "bmf文件格式转换错误！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    End If
 
End Function
