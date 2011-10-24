VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLib 
   Caption         =   "封装库"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "MakeLib"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6795
   ScaleWidth      =   9195
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "帮助"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5535
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "更新"
      Height          =   615
      Left            =   480
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "版本信息"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.Menu menuMountTpye 
      Caption         =   "封装类型"
      Visible         =   0   'False
      Begin VB.Menu menuTypeSub 
         Caption         =   "S：贴装元件"
         Index           =   0
      End
      Begin VB.Menu menuTypeSub 
         Caption         =   "L：插装元件"
         Index           =   1
      End
      Begin VB.Menu menuTypeSub 
         Caption         =   "S+：贴装元件"
         Index           =   2
      End
      Begin VB.Menu menuTypeSub 
         Caption         =   "N：None元件"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'分类库
Public Enum MenuType
MENU_S = 0   'SMD Type
MENU_L       'LEAD Type
MENU_SP      'SMD Type
MENU_N       'None Type
End Enum

Private Sub UpdateSTD(newlibfile As String)
    '读取旧库文件 新库文件 将旧库文件相同封装的的贴装信息转移到新库文件中
    '最后打开新库文件提示填入新的封装对应的贴装类型
    
    If newlibfile = "" Then
        Exit Sub
    End If
    
    Dim oldLibLine()          As String
    Dim newLibLine()          As String
    Dim tmpLibLine()          As String
    
    Dim oldAtom()             As String
    Dim newAtom()             As String
    
    oldLibLine = OpenLibs()
    
    newLibLine = Split(GetFileContents(newlibfile), vbCrLf)
    tmpLibLine = newLibLine
    
    '添加版本信息 等指定的信息
    newAtom = Split(tmpLibLine(0), Space(1))
    If UBound(newAtom) = 13 Then
        newLibLine(0) = newLibLine(0) & Space(14) & "VERSION:" & Space(2) & CStr(Now)
    End If
    
    newAtom = Split(tmpLibLine(1), Space(1))
    If UBound(newAtom) = 34 Then
        newLibLine(1) = newLibLine(1) & Space(14) & "Mount Type"
    End If
    'newLibLine(2) = newLibLine(2) 获取文件内容的时候去掉了空行
    
    '将旧库的信息导入到新库 遍历旧库
    Dim i As Integer
    Dim j As Integer
    
    For i = 2 To UBound(tmpLibLine) - 1
        Do While InStr(tmpLibLine(i), Space(2))
            tmpLibLine(i) = Replace(tmpLibLine(i), Space(2), Space(1)) '清除多余的空格
        Loop
    Next i
    
    On Error GoTo ErrorHandle
    For i = 2 To UBound(oldLibLine) - 1
        oldAtom = Split(oldLibLine(i), Space(1))
            For j = 2 To UBound(tmpLibLine) - 1
                newAtom = Split(tmpLibLine(j), Space(1))
                If UBound(newAtom) = 2 And newAtom(0) = oldAtom(0) Then
                     tmpLibLine(j) = tmpLibLine(j) & Space(1) & oldAtom(3)
                     newLibLine(j) = newLibLine(j) & Space(10) & oldAtom(3)
                End If
            Next j
    Next i
    
    '检查是否所有的封装都有对应的贴装类型，没有的话提示添加
    'For j = 2 To UBound(tmpLibLine) - 1
    '    newAtom = Split(tmpLibLine(j), Space(1))
    '    If UBound(newAtom) = 3 Then
    '
    '    ElseIf UBound(newAtom) = 2 Then
    '        '添加贴装类型
    '        MsgBox "请手动打开库文件填写未完成的封装分类!", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
    '        Exit For
    '    Else
    '        MsgBox "导入的库文件是错误的，拒绝更新库文件！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
    '        Exit Sub
    '    End If
    'Next j
    
    If Dir(LibFilePath) <> "" Then
        Kill LibFilePath
    End If
    
    Open LibFilePath For Binary Access Write As #1
    Seek #1, 1
    Put #1, , newLibLine(0) & vbCrLf
    Put #1, , newLibLine(1) & vbCrLf
    Put #1, , vbCrLf
    
    For j = 2 To UBound(newLibLine) - 1
        Put #1, , newLibLine(j) & vbCrLf
    Next j
        
    Put #1, , vbCrLf

    Close #1
        
    'MsgBox "Done!", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
    '打开库文件
    'Shell "notepad " & LibFilePath, vbMaximizedFocus
    
    LoadLibs
    
ErrorHandle:

End Sub


Private Sub Command1_Click()
    Dim filePath As String
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "请选择STD.LST文件"
    CommonDialog1.Filter = "All File(*.*)|*.*|lst files(*.lst)|*.lst"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    
    filePath = CommonDialog1.FileName
    
    Dim filetype As String
    filetype = Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)
    
    If filetype = ".lst" Or filetype = ".LST" Then
        UpdateSTD filePath
    End If
End Sub

Private Sub Command1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    '允许拖放操作
    Dim PLResultFile As Variant
    Dim filePath As String
    
    On Error Resume Next
    For Each PLResultFile In Data.Files
        filePath = PLResultFile
    Next
    
    Dim filetype As String
    filetype = Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)
    
    If filetype = ".lst" Or filetype = ".LST" Then
        UpdateSTD filePath
    End If
End Sub

Private Sub Command2_Click()
    Dim msgstr As String

    msgstr = msgstr + "第1步：选择从Power PCB中导出的STD.lst。" & vbCrLf & vbCrLf
    msgstr = msgstr + "第2步：在需要修改的封装上使用鼠标点击选择正确的封装类型。" & vbCrLf & vbCrLf
    
    msgstr = msgstr + "注意：" & vbCrLf & vbCrLf
    msgstr = msgstr + "最终修改后的库文件需提交到版本库中。" & vbCrLf
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "帮助信息"
End Sub

Private Sub Form_Load()
    SetWindowsPos_TopMost Me.hwnd
    
    Label1.Caption = "封装库版本：" & GetLibsVersion
    
    '读取旧库文件 显示在ListView上
    InitList
    
End Sub

Private Sub InitList()

    '添加库文件内容到ListView
    Dim gwidth        As Integer
    
    '初始化ListView1
    gwidth = ListView1.Width / 40
    
    ListView1.ColumnHeaders.Add , , "编号", gwidth * 3
    ListView1.ColumnHeaders.Add , , "封装描述", gwidth * 18
    ListView1.ColumnHeaders.Add , , "封装类型", gwidth * 6
    ListView1.ColumnHeaders.Add , , "修改时间", gwidth * 10
    
    LoadLibs
    
End Sub

Private Sub LoadLibs()
    Dim LibLine()     As String
    Dim LibAtom()     As String
    Dim i             As Integer
    
    '清空列表
    ListView1.ListItems.Clear
    
    '添加元素
    LibLine = OpenLibs()
    
    For i = 2 To UBound(LibLine) - 1
        LibAtom = Split(LibLine(i), Space(1))
        ListView1.ListItems.Add , , (i - 1) & ""
        If UBound(LibAtom) = 3 Then
            ListView1.ListItems(i - 1).SubItems(2) = LibAtom(3)
        End If
            
        ListView1.ListItems(i - 1).SubItems(1) = LibAtom(0)
        ListView1.ListItems(i - 1).SubItems(3) = LibAtom(1) & Space(1) & LibAtom(2)
    Next
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Command1_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '检查文件信息是否完整，不完整需要提示手动修改
    
    '检验库文件的完整性和正确性
    Dim LibLine()          As String
    Dim LibAtom()          As String
    Dim i                  As Integer
    
    LibLine = OpenLibs()
    For i = 2 To UBound(LibLine) - 1
        LibAtom = Split(LibLine(i), Space(1))
        If UBound(LibAtom) <> 3 Then
            MsgBox "库文件STD.lst不完整或错误！请手动修改STD.lst库文件！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            End
        End If
    Next i
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Dim j As Long, i As Long
    
    '响应鼠标左键 和 右键事件 和 中键事件
    If Button = vbLeftButton Or Button = vbRightButton Or Button = vbMiddleButton Then
        If ListView1.HitTest(x, y) Is Nothing Then
            Exit Sub
        End If
        
        j = ListView1.HitTest(x, y).Index
        ListView1.ListItems(j).Selected = True
        
        For i = 0 To 3
            menuTypeSub(i).Checked = False
        Next
        
        Select Case ListView1.SelectedItem.SubItems(2)
            Case "S": menuTypeSub(MENU_S).Checked = True
            Case "L": menuTypeSub(MENU_L).Checked = True
            Case "S+": menuTypeSub(MENU_SP).Checked = True
            Case "N": menuTypeSub(MENU_N).Checked = True
        End Select
        
        PopupMenu menuMountTpye
    End If
    
    If Button = vbLeftButton Then
        '不响应鼠标左键
    End If
    
End Sub

Private Sub menuTypeSub_Click(Index As Integer)
    On Error Resume Next
    Dim ModLibLine  As String
    Dim ModAtom()   As String

    If menuTypeSub(Index).Checked = True Then Exit Sub

    Select Case Index
        Case MENU_S
        ListView1.SelectedItem.SubItems(2) = "S"
        Case MENU_L
        ListView1.SelectedItem.SubItems(2) = "L"
        Case MENU_SP
        ListView1.SelectedItem.SubItems(2) = "S+"
        Case MENU_N
        ListView1.SelectedItem.SubItems(2) = "N"
    End Select

    '将相应的修改同步到库文件中
    If ListView1.SelectedItem.SubItems(2) <> "" Then
        ModLibLine = ReadTxt(LibFilePath, Val(ListView1.SelectedItem.Text) + 3)
        ModAtom = Split(ModLibLine, Space$(1))
        Select Case ModAtom(UBound(ModAtom))
            Case "S", "L", "N"
            ModLibLine = Left(ModLibLine, Len(ModLibLine) - 1) + ListView1.SelectedItem.SubItems(2)
            Case "S+"
            ModLibLine = Left(ModLibLine, Len(ModLibLine) - 2) + ListView1.SelectedItem.SubItems(2)
            Case Else
            ModLibLine = ModLibLine + Space(10) + ListView1.SelectedItem.SubItems(2)
        End Select

        WriteTxt LibFilePath, Val(ListView1.SelectedItem.Text) + 3, ModLibLine
    End If

End Sub

'参数一 要写入的文件地址，参数二 修改的行数 ，参数三 写入或替换的字符串
Public Function WriteTxt(strSourceFile As String, intRow As Long, StrLineNew As String)

    Dim StrOut As String, tmpStrLine As String
    Dim x As Long
    If Dir(strSourceFile) <> "" Then
        Open strSourceFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpStrLine
            x = x + 1
            If x = intRow Then tmpStrLine = StrLineNew
            StrOut = StrOut & tmpStrLine & vbCrLf
            'Debug.Print x
        Loop
        Close #1
    Else
        StrOut = StrLineNew
    End If
    
    '多了一个换行符？
    Open strSourceFile For Output As #1
    Print #1, StrOut
    Close #1

End Function

'返回 要输出的文本，参数一 文件地址，参数二 读取的行数
Public Function ReadTxt(StrFile As String, intRow As Long) As String
    Dim StrOut As String, tmpStrLine As String
    Dim x As Long
    
    If Dir(StrFile, vbNormal) <> "" Then
        Open StrFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpStrLine
            x = x + 1
            If x = intRow Then ReadTxt = tmpStrLine: Exit Do
        Loop
        Close #1
    End If
    
End Function

'ListView的默认的排序功能都是按照字符串顺序排的，那样对数字顺序时，如果升序排列，9会排在10的后面。
'以下程序将修正这一问题
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    
    '计时器 调试测试时使用
    'Dim lngStart As Long
    'lngStart = GetTickCount
    
    With ListView1
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        LockWindowUpdate .hwnd
        
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = ColumnHeader.Index - 1
        
        If ColumnHeader.Text = "编号" Then
            
            strFormat = String(30, "0") & "." & String(30, "0")
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), strFormat)
                                Else
                                    .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), strFormat)
                                Else
                                    .Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            .SortOrder = .SortOrder Xor 1
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
       Else
            .SortOrder = .SortOrder Xor 1
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
                
        End If
            
        LockWindowUpdate 0&
        
        .MousePointer = lngCursor
        
    End With
    
    'MsgBox "行数: " & ListView1.ListItems.Count & " 用时: " & GetTickCount - lngStart & "ms"
    
End Sub

Private Function InvNumber(ByVal Number As String) As String
    Static i As Integer
    For i = 1 To Len(Number)
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function
