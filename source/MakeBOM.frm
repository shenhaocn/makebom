VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOM生成工具"
   ClientHeight    =   2070
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MakeBOM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2070
   ScaleWidth      =   4620
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "库存类型"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1275
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "MakeBOM.frx":628A
         Left            =   60
         List            =   "MakeBOM.frx":6297
         TabIndex        =   7
         Text            =   "TP1"
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "机型名称"
      Height          =   615
      Left            =   300
      TabIndex        =   4
      Top             =   120
      Width           =   2475
      Begin VB.TextBox ItemNameText 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "tsv文件编码"
      Height          =   615
      Left            =   300
      TabIndex        =   2
      Top             =   840
      Width           =   1395
      Begin VB.ComboBox UTFCombo 
         Height          =   300
         ItemData        =   "MakeBOM.frx":62AA
         Left            =   60
         List            =   "MakeBOM.frx":62BA
         TabIndex        =   3
         Top             =   240
         Width           =   1275
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   1755
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6703
            MinWidth        =   6703
            Text            =   "请选择.BOM文件..."
            TextSave        =   "请选择.BOM文件..."
            Key             =   "status_text"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "0%"
            TextSave        =   "0%"
            Key             =   "process_text"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton MakeBOM_Command 
      Caption         =   "生成BOM"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   840
      Width           =   2355
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -240
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu menu_null 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu menu_lib 
      Caption         =   "分类封装库"
   End
   Begin VB.Menu menu_update 
      Caption         =   "封装库更新"
   End
   Begin VB.Menu menu_feedback 
      Caption         =   "反馈"
   End
   Begin VB.Menu menu_about 
      Caption         =   "关于"
   End
   Begin VB.Menu menu_winpos 
      Caption         =   "--"
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'=========================================================
'窗口主体实现文件，仅包含窗口功能实现代码以及执行流程控制
'=========================================================

Private Sub Form_Load()
    
    '初始化数据库
    If InitLib(App.Path & "\STD.lst") = False Then
        MakeBOM_Command.Enabled = False
    End If
    
    '获取上次工作目录
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "上次工作目录", "E:\")
    ItemName = GetSetting(App.EXEName, "TaskName", "上一次项目机型名称", "")
    
    ItemNameText.Text = ItemName
    
    '获取程序设置
    UTFCombo.Text = GetSetting(App.EXEName, "tsvEncoder", "tsv文件编码", "UTF-8")
    Combo1.Text = GetSetting(App.EXEName, "SelectStorage", "库存类型", "TP1")
    
    '初始化窗口位置和状态
    Dim X As String
    Dim Y As String
    X = GetSetting(App.EXEName, "WindowPosition", "Left")
    Y = GetSetting(App.EXEName, "WindowPosition", "Top")
        
    If X <> "" Then
        Me.Move X, Y
    End If
    
    Dim AlwaysOnTop As String
    AlwaysOnTop = GetSetting(App.EXEName, "WindowPosition", "AlwaysOnTop", "--")
    If AlwaysOnTop = "--" Then
        menu_winpos.Caption = "--"
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '直接调用按钮的拖放效果
    MakeBOM_Command_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '程序配置数据
    SaveSetting App.EXEName, "tsvEncoder", "tsv文件编码", UTFCombo.Text
    SaveSetting App.EXEName, "SelectStorage", "库存类型", Combo1.Text
    
    SaveSetting App.EXEName, "ProjectDir", "上次工作目录", ProjectDir
    SaveSetting App.EXEName, "TaskName", "上一次项目机型名称", ItemNameText.Text
    
    '窗口位置
    SaveSetting App.EXEName, "WindowPosition", "Left", Me.Left
    SaveSetting App.EXEName, "WindowPosition", "Top", Me.Top
    SaveSetting App.EXEName, "WindowPosition", "AlwaysOnTop", menu_winpos.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '结束所有的窗体
    Dim counter As Integer
    For counter = 0 To Forms.Count - 1
        Unload Forms(counter)
    Next
    
    End
    
End Sub

Private Sub menu_about_Click()
    frmAbout.Show 1
End Sub

Private Sub menu_feedback_Click()
    ShellExecute 0, "open", "mailto:shenhao@tp-link.net?subject=【MakeBOM】反馈&Body=", "", "", 1
End Sub

Private Sub menu_lib_Click()
    '打开库文件
    Shell "notepad " & LibFilePath, vbMaximizedFocus
End Sub

Private Sub menu_update_Click()
    frmUpdateLib.Show 1
End Sub

Private Sub menu_winpos_Click()
    '设置窗口是否固定在最上层
    If menu_winpos.Caption = "|" Then
        menu_winpos.Caption = "--"
        '将 窗口设定成永远保持在最上层
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        '取消最上层设定
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
End Sub

Private Sub UTFCombo_Click()
    SaveSetting App.EXEName, "tsvEncoder", "tsv文件编码", UTFCombo.Text
End Sub

Private Sub UTFCombo_LostFocus()
    Select Case UTFCombo.Text
    Case "UTF-8"
            
    Case "ANSI"
        
    Case "UTF-16LE"
        
    Case "UTF-16BE"
        
    Case Else
        MsgBox "不支持的文本文件编码！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
        UTFCombo.Text = GetSetting(App.EXEName, "tsvEncoder", "tsv文件编码", "UTF-8")
    End Select
    
    SaveSetting App.EXEName, "tsvEncoder", "tsv文件编码", UTFCombo.Text
End Sub

Private Sub Combo1_Click()
    SaveSetting App.EXEName, "SelectStorage", "库存类型", Combo1.Text
End Sub

Private Sub Combo1_LostFocus()
    Select Case Combo1.Text
    Case "TP1"

    Case "TP2"

    Case "TP3"

    Case Else
        MsgBox "不支持的库存类型！请重新选择", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
        Combo1.Text = "TP1"
    End Select
    
    SaveSetting App.EXEName, "SelectStorage", "库存类型", Combo1.Text
End Sub


Private Sub MakeBOM_Command_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '允许拖放操作
    Dim PLResultFile As Variant
    Dim filePath As String
    
    On Error Resume Next
    For Each PLResultFile In Data.Files
        filePath = PLResultFile
    Next
    
    Dim filetype As String
    filetype = Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)
    
    If filetype = ".BOM" Or filetype = ".bom" Then
        BuildProjectPath filePath
        
        BomStage_One
        Exit Sub
    End If
    
    If filetype = ".tsv" Then
        tsvFilePath = filePath
        
        If BomFilePath = "" Then
            MsgBox "请先选择.BOM文件所在路径！", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
            Exit Sub
        End If
        
        BomStage_Two
    End If
End Sub

Private Sub MakeBOM_Command_Click()
    Dim GetPath As String
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "上次工作目录", "E:\")
    
    If BomFilePath = "" Then
        CommonDialog1.InitDir = ProjectDir
        CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "请选择.BOM文件"
        CommonDialog1.Filter = "All File(*.*)|*.*|BOM files(*.BOM)|*.BOM"
        CommonDialog1.FilterIndex = 2
        CommonDialog1.ShowOpen
        
        GetPath = CommonDialog1.FileName
        
        If GetPath = "" Then
            Exit Sub
        End If
        
        Dim isbom As String
        isbom = Right(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)
    
        If isbom <> ".BOM" And isbom <> ".bom" Then
            MsgBox "必须为.BOM文件！", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "警告"
            ClearPath
            Exit Sub
        Else
            BuildProjectPath GetPath
            
            BomStage_One
        End If
    Else
        If tsvFilePath = "" Then
            CommonDialog1.FileName = ""
            CommonDialog1.DialogTitle = "请选择.tsv文件"
            CommonDialog1.Filter = "All File(*.*)|*.*|tsv files(*.tsv)|*.tsv"
            CommonDialog1.FilterIndex = 2
            CommonDialog1.ShowOpen
            
            GetPath = CommonDialog1.FileName
            
            If GetPath = "" Then
                Exit Sub
            End If
            
            Dim istsv As String
            istsv = Right(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)
            
            If istsv <> ".tsv" Then
                MsgBox "必须为.tsv文件！", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "警告"
                Exit Sub
            Else
                tsvFilePath = GetPath
            End If
        End If
    
        BomStage_Two
    End If
End Sub


Private Sub BomStage_One()

    Dim msgstr As String
    
    '删除过时的文件
    msgstr = SaveAsPath & "_PCBA_BOM.xls" & vbCrLf
    msgstr = msgstr + SaveAsPath & "_批量资源查询.xls" & vbCrLf
    msgstr = msgstr + SaveAsPath & "_None_PartRef.xls" & vbCrLf
    msgstr = msgstr + SaveAsPath & "_NC_DBG.xls" & vbCrLf & vbCrLf
    msgstr = msgstr + "已经存在，使用这些文件？" & vbCrLf
    
    If Dir(SaveAsPath & "_PCBA_BOM.xls") <> "" _
            Or Dir(SaveAsPath & "_批量资源查询.xls") <> "" _
                Or Dir(SaveAsPath & "_None_PartRef.xls") <> "" _
                Or Dir(SaveAsPath & "_NC_DBG.xls") <> "" Then
        If MsgBox(msgstr, vbInformation + vbMsgBoxSetForeground + vbYesNo) = vbYes Then
            Process 10, "读取元件数..."
            CalcPartNum
            
            Process 20, "开始生成BOM第2阶段，选择tsv文件路径..."
            If tsvFilePath = "" Then
                MakeBOM_Command_Click
            Else
                BomStage_Two
            End If
            
            Exit Sub
        Else
            KillExcel SaveAsPath & "_PCBA_BOM.xls"
            KillExcel SaveAsPath & "_批量资源查询.xls"
            KillExcel SaveAsPath & "_None_PartRef.xls"
            KillExcel SaveAsPath & "_NC_DBG.xls"
        End If
    End If
    
    Process 2, "读取.BOM文件信息 ..."
    '读取.BOM文件信息
    If ReadBomFile = False Then
        Process 100, ".BOM文件信息 ...读取错误！"
        ClearPath
        Exit Sub
    End If
    
    '由模板创建Excel文件
    ExcelCreate
    
    '填充来自orCAD BOM的数据
    BomDraft
    
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    If MsgBox("是否打开ERP系统进行批量查询？", vbQuestion + vbMsgBoxSetForeground + vbYesNo) = vbYes Then
        AutoLoginERP "RD_ENGINEER", "123456"
        'FindERP
    End If
    
End Sub


Public Sub BomStage_Two()
    
    Dim msgstr As String
    '删除过时的文件
    'msgstr = SaveAsPath & "_领料BOM.xls" & vbCrLf
    'msgstr = msgstr + SaveAsPath & "_生产BOM.xls" & vbCrLf
    'msgstr = msgstr + SaveAsPath & "_调试BOM.xls" & vbCrLf & vbCrLf
    'msgstr = msgstr + "已经存在，是否删除以便重新生成？" & vbCrLf
    
    If Dir(SaveAsPath & "_领料BOM.xls") <> "" _
            Or Dir(SaveAsPath & "_生产BOM.xls") <> "" _
                Or Dir(SaveAsPath & "_调试BOM.xls") <> "" Then
        'If MsgBox(msgstr, vbInformation + vbMsgBoxSetForeground + vbYesNo) = vbYes Then
        KillExcel SaveAsPath & "_领料BOM.xls"
        KillExcel SaveAsPath & "_生产BOM.xls"
        KillExcel SaveAsPath & "_调试BOM.xls"
        'Else
        '    Exit Sub
        'End If
    End If
    
    '输出所有BOM
    If CreateAllBOM = False Then
        Process 100, "创建BOM文件失败！"
        GoTo ErrorHandle
    End If
    '调整部分BOM格式或数据
    If BomAdjust = False Then
        Process 100, "调整BOM文件失败！"
        GoTo ErrorHandle
    End If
    
    '输出完整BOM
    If ImportTSV(SaveAsPath & "_PCBA_BOM.xls", 80) = False Then
        Process 100, "tsv文件错误！请更新tsv文件！"
        GoTo ErrorHandle
    End If
    
    If ImportTSV(SaveAsPath & "_领料BOM.xls", 84) = False Then
        Process 100, "tsv文件错误！请更新tsv文件！"
        GoTo ErrorHandle
    End If
    
    If ImportTSV(SaveAsPath & "_调试BOM.xls", 88) = False Then
        Process 100, "tsv文件错误！请更新tsv文件！"
        GoTo ErrorHandle
    End If
    
    If ImportTSV(SaveAsPath & "_生产BOM.xls", 92) = False Then
        Process 100, "tsv文件错误！请更新tsv文件！"
        GoTo ErrorHandle
    End If
    
    '打开生成的BOM，以备检查
    'ShellExecute 0, "open", SaveAsPath & "_PCBA_BOM.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_NC_DBG.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_None_PartRef.xls", "", "", 1
    
    'ShellExecute 0, "open", SaveAsPath & "_领料BOM.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_调试BOM.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_生产BOM.xls", "", "", 1
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    '清空程序依赖的路径信息，便于下次转换开始
    ClearPath
    
    Process 100, "完成！"
    Exit Sub

ErrorHandle:
    
    '删除未能成功生成的文件
    KillExcel SaveAsPath & "_领料BOM.xls"
    KillExcel SaveAsPath & "_生产BOM.xls"
    KillExcel SaveAsPath & "_调试BOM.xls"
    
    ClearPath

End Sub

