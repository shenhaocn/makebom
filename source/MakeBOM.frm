VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOM生成工具"
   ClientHeight    =   4140
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4545
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
   ScaleHeight     =   4140
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "生成BOM"
      Height          =   1035
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   4155
      Begin VB.CheckBox CheckNcDbg 
         Caption         =   "NC DBG元件"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   300
         Width           =   1395
      End
      Begin VB.CheckBox CheckAll 
         Caption         =   "全选"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox Check_生产 
         Caption         =   "生产BOM"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox Check_调试 
         Caption         =   "调试BOM"
         Height          =   255
         Left            =   1380
         TabIndex        =   9
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox Check_领料 
         Caption         =   "领料BOM"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox CheckPreBom 
         Caption         =   "预BOM"
         Height          =   255
         Left            =   1380
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "库存类型"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "MakeBOM.frx":628A
         Left            =   60
         List            =   "MakeBOM.frx":6297
         TabIndex        =   5
         Text            =   "TP1"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "机型名称"
      Height          =   615
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   2775
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
         Height          =   285
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3825
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6350
            MinWidth        =   6350
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "status_text"
            Object.ToolTipText     =   "程序运行状态描述"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1693
            MinWidth        =   1693
            Text            =   "0%"
            TextSave        =   "0%"
            Key             =   "process_text"
            Object.ToolTipText     =   "执行进度"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command_ImportBom 
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
      Height          =   1335
      Left            =   180
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   2100
      Width           =   4155
   End
   Begin VB.Menu menu_lib 
      Caption         =   "封装库"
   End
   Begin VB.Menu menu_feedback 
      Caption         =   "反馈"
      Visible         =   0   'False
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
'*************************************************************************
'**模 块 名：MainForm
'**说    明：TP-LINK SMB Switch Product Line Hardware Group
'**          版权所有2011 - 2012(C)
'**
'**创 建 人：Shenhao
'**日    期：2011-10-22 12:08:02
'**修 改 人：
'**日    期：
'**描    述：窗口主体实现文件，仅包含窗口功能实现代码以及执行流程控制
'**版    本：V3.2.38
'*************************************************************************

Option Explicit

Private Sub Form_Load()
    
    '初始化数据库
    If InitLib(App.Path & "\STD.lst") = False Then
        Command_ImportBom.Enabled = False
    End If
    
    '获取上次工作目录
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "上次工作目录", "E:\")
    ItemName = GetSetting(App.EXEName, "TaskName", "上一次项目机型名称", "")
    
    ItemNameText.Text = ItemName
    
    '获取程序设置
    Combo1.Text = GetSetting(App.EXEName, "SelectStorage", "库存类型", "TP1")
    
    '初始化窗口位置和状态
    Dim X As String
    Dim Y As String
    X = GetSetting(App.EXEName, "WindowPosition", "Left")
    Y = GetSetting(App.EXEName, "WindowPosition", "Top")
        
    If X <> "" And Y <> "" And _
       Val(X) < Screen.Width And Val(Y) < Screen.Height And _
       Val(X) > 0 And Val(Y) > 0 Then
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
    Command_ImportBom_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '程序配置数据
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
    'Shell "notepad " & LibFilePath, vbMaximizedFocus
    frmLib.Show 1
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

Private Sub CheckAll_Click()

    If CheckAll.Value = Checked Then
        CheckPreBom.Value = Checked
        CheckNcDbg.Value = Checked
        Check_领料.Value = Checked
        Check_调试.Value = Checked
        Check_生产.Value = Checked
    End If

End Sub

Private Sub CheckCheck()
    If CheckPreBom.Value = Checked And _
       CheckNcDbg.Value = Checked And _
       Check_领料.Value = Checked And _
       Check_调试.Value = Checked And _
       Check_生产.Value = Checked Then
       
        CheckAll.Value = Checked
    Else
        CheckAll.Value = Unchecked
    End If
End Sub

Private Sub Check_调试_Click()
    CheckCheck
End Sub

Private Sub Check_领料_Click()
    CheckCheck
End Sub

Private Sub Check_生产_Click()
    CheckCheck
End Sub

Private Sub CheckNcDbg_Click()
    CheckCheck
End Sub

Private Sub CheckPreBom_Click()
    CheckCheck
End Sub

Private Sub Command_ImportBom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
        
        BomStage_One
    End If
End Sub

Private Sub Command_ImportBom_Click()
    Dim GetPath As String
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "上次工作目录", "E:\")
    
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
    End If
    
    BuildProjectPath GetPath

    BomStage_One
    
End Sub

Private Sub BomStage_One()
    
    '==============================================
    '本阶段将会生成BOM Maker File
    '1.导入tsv文件信息到bmf文件中
    '2.保留曾经有的部分信息 如描述 库存等
    '3.将所需信息整理为标准格式 便于后一阶段读取
    '4.使用文本格式便于版本控制
    '==============================================
    
    Dim GetPath As String
    
    KillBom

    Process 2, "读取.BOM文件信息 ..."
    '读取.BOM文件信息
    If ReadBomFile = False Then
        Process 100, ".BOM文件信息 ...读取错误！"
        ClearPath
        Exit Sub
    End If
    
    '创建批量查询文件
    BomMakePLExcel
    
    '填充来自orCAD BOM的数据并且创建新的.bmf文件
    BmfMaker
    
    '默认tsv文件在工作目录下
    tsvFilePath = ProjectDir + "fnd_gfm.tsv"
    
    '查看.BOM目录下是否有tsv文件，有的话直接导入 没有就询问是否进入ERP查询
    If Dir(tsvFilePath) = "" Then
    
        Dim resultL As VbMsgBoxResult
        resultL = MsgBox("在工作目录下未找到合法的批量查询结果文件！" & vbCrLf & vbCrLf & vbCrLf & _
                  "是否打开ERP系统进行批量查询？" & vbCrLf & vbCrLf & _
                  "是-登录ERP系统开始查询" & vbCrLf & vbCrLf & _
                  "否-选择TSV文件路径" & vbCrLf, _
                  vbQuestion + vbMsgBoxSetForeground + vbYesNoCancel)
                  
        If resultL = vbYes Then
                  
            AutoLoginERP "RD_ENGINEER", "123456"
            'FindERP
            Exit Sub
            
        ElseIf resultL = vbNo Then
        
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
            
        ElseIf resultL = vbCancel Then
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    '导入tsv文件内信息
    ImportTSV
    
    '转换BMF文件格式避免出现乱码
    BmfToAnsi
    
    '查找bmf文件中有料号，但是没有物料描述的行
    '给出提示是否自动联网更新物料描述
    'GetInfoFromERP
    'GetInfoFromERP "RD_ENGINEER", "123456"
    
    '直接进入第2阶段 生成Excel格式BOM阶段
    BomStage_Two
    
End Sub


'*************************************************************************
'**函 数 名：BomStage_Two
'**输    入：无
'**输    出：无
'**功能描述：根据CheckBox的状态创建Excel文件和生成相应的BOM
'            流程如下:
'            1.根据模版创建Excel BOM
'            2.根据需要调整Excel 格式
'            3.读取bmf(BOM Maker File)文件 将信息填入Excel
'            4.根据信息调整Excel格式
'            5.扫描Excel格式 修正部分格式
'            6.完成
'            注意：领料BOM中的库存信息必须保证是最新的。
'                  因此程序会检查tsv文件的产生时间
'                  时间不在三天内会提示，重新查询�
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-22 12:11:09
'**修 改 人：
'**日    期：
'**版    本：V3.2.38
'*************************************************************************
Public Sub BomStage_Two()
    
    '由模板创建Excel文件 并生成需要的Excel BOM文件
    BomCreate
    
     '============================================
    '显示结果信息 元件数量个数
    '============================================
    'PartNum(0) : NcPartNum
    'PartNum(1) : DbgPartNum
    'PartNum(2) : DbNcPartNum
    
    'PartNum(3) : LeadPartNum
    'PartNum(4) : SmtPartNum
    'PartNum(5) : OtherPartNum
    '============================================
    
    Dim msgstr As String
    msgstr = "             BOM 文件创建成功！" & vbCrLf & vbCrLf
    msgstr = msgstr + "          插装   元件个数为   ： " & PartNum(3) & vbCrLf
    msgstr = msgstr + "          贴装   元件个数为   ： " & PartNum(4) & vbCrLf
    msgstr = msgstr + "          其他   元件个数为   ： " & PartNum(5) & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     元件个数为   ： " & PartNum(0) & vbCrLf
    msgstr = msgstr + "          DBG    元件个数为   ： " & PartNum(1) & vbCrLf
    msgstr = msgstr + "          DBG_NC 元件个数为   ： " & PartNum(2) & vbCrLf & vbCrLf
    msgstr = msgstr + "          生成的bmf文件不建议手动修改" & vbCrLf & vbCrLf
    msgstr = msgstr + "    注意：生成的BOM文件需要检查修改后才可供评审 "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "BOM 文件创建成功"
    
    '打开生成的BOM目录
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    '清空程序依赖的路径信息，便于下次转换开始
    ClearPath
    
    Process 100, "完成！"
    
    Exit Sub

ErrorHandle:
    
    '删除未能成功生成的文件
    KillBom
    
    ClearPath

End Sub


Private Sub BomCreate()
    If CheckPreBom.Value = Checked Then
        ExcelCreate BOM_预
        CreateBOM BOM_预
        
    End If
    
    If CheckNcDbg.Value = Checked Then
        ExcelCreate BOM_NCDBG
        CreateBOM BOM_NCDBG
        
        ExcelCreate BOM_NONE
        CreateBOM BOM_NONE
        
    End If
    
    If Check_领料.Value = Checked Then
        'tsv文件是否失效？
        If DateDiff("d", CDate(GetFileWriteTime(tsvFilePath)), Now) > 3 Then
            MsgBox "tsv文件已经过期[" & GetFileWriteTime(tsvFilePath) & "]！" & vbCrLf & vbCrLf & _
                   "生成领料BOM需最新的tsv文件！", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "警告"
        Else
            ExcelCreate BOM_领料
            CreateBOM BOM_领料
        End If
        
    End If
    
    If Check_调试.Value = Checked Then
        ExcelCreate BOM_调试
        CreateBOM BOM_调试
        
    End If
    
    If Check_生产.Value = Checked Then
        ExcelCreate BOM_生产
        CreateBOM BOM_生产
        
    End If
End Sub

Private Sub KillBom()

    KillExcel SaveAsPath & "_预BOM_BMF.xls"
    KillExcel SaveAsPath & "_批量资源查询.xls"
    KillExcel SaveAsPath & "_None_PartRef.xls"
    KillExcel SaveAsPath & "_NC_DBG.xls"
    
    KillExcel SaveAsPath & "_领料BOM.xls"
    KillExcel SaveAsPath & "_生产BOM.xls"
    KillExcel SaveAsPath & "_调试BOM.xls"
    
End Sub

