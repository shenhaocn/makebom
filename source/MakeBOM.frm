VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MakeBOM(BOM生成工具)"
   ClientHeight    =   4335
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
   ScaleHeight     =   4335
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "BOM类型"
      Height          =   1395
      Left            =   180
      TabIndex        =   6
      Top             =   735
      Width           =   4155
      Begin VB.CheckBox CheckNcDbg 
         Caption         =   "NCDBG元件"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox CheckNone 
         Caption         =   "None元件"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox Check_生产 
         Caption         =   "生产BOM"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check_调试 
         Caption         =   "调试BOM"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox Check_领料 
         Caption         =   "领料BOM"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1155
      End
      Begin VB.CheckBox CheckPreBom 
         Caption         =   "预BOM"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "清空"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "反选"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "全选"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   855
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
      Top             =   4020
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
      Height          =   1455
      Left            =   180
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   2280
      Width           =   4155
   End
   Begin VB.Menu menu_lib 
      Caption         =   "封装库"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu menu_feedback 
      Caption         =   "反馈"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
   End
   Begin VB.Menu menu_about 
      Caption         =   "关于"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu menu_winpos 
      Caption         =   "--"
      NegotiatePosition=   2  'Middle
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

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Form_Initialize()
    InitCommonControls
End Sub

'载入程序配置
Private Sub Form_Load()
    
    Dim X As Long
    Dim Y As Long
    
    '初始化数据库
    If InitLib(App.Path & "\STD.lst") = False Then
        Command_ImportBom.Enabled = False
    End If
    
    '获取程序设置
    ItemName = GetRegValue(App.EXEName, "Product", "")
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    ItemNameText.Text = ItemName
    Combo1.Text = GetRegValue(App.EXEName, "Storage", "TP1")
    
    '初始化窗口位置 默认在屏幕中央
    X = GetRegValue(App.EXEName, "WinLeft", Screen.Width / 2 - Me.Width / 2)
    Y = GetRegValue(App.EXEName, "WinTop", Screen.Height / 2 - Me.Height / 2)
        
    If X > Screen.Width Or Y > Screen.Height Or _
       X < 0 Or Y < 0 Then
        X = Screen.Width / 2 - Me.Width / 2
        Y = Screen.Height / 2 - Me.Height / 2
    End If
    
    Me.Move X, Y
    
    '获取窗口状态
    If GetRegValue(App.EXEName, "OnTop", 1) = 1 Then
        menu_winpos.Caption = "--"
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
    Command_ImportBom.Caption = "生成BOM" & vbCrLf & vbCrLf & "（BomChecker）"
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '直接调用按钮的拖放效果
    Command_ImportBom_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '程序配置数据
    SetRegValue App.EXEName, "Storage", iREG_SZ, Combo1.Text
    
    SetRegValue App.EXEName, "ProjectDir", iREG_SZ, ProjectDir
    SetRegValue App.EXEName, "Product", iREG_SZ, ItemNameText.Text
    
    '窗口位置
    SetRegValue App.EXEName, "WinLeft", iREG_DWORD, Me.Left
    SetRegValue App.EXEName, "WinTop", iREG_DWORD, Me.Top
    
    '窗口状态
    If menu_winpos.Caption = "|" Then
        SetRegValue App.EXEName, "OnTop", iREG_DWORD, 0
    Else
        SetRegValue App.EXEName, "OnTop", iREG_DWORD, 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '结束所有的窗体
    Dim counter As Integer
    For counter = 0 To Forms.Count - 1
        Unload Forms(counter)
    Next
    
    End
    
End Sub

'菜单
Private Sub menu_lib_Click()
    '打开库文件
    frmLib.Show 1
End Sub

Private Sub menu_about_Click()
    frmAbout.Show 1
End Sub

Private Sub menu_feedback_Click()
    ShellExecute 0, "open", "mailto:shenhao@tp-link.net?subject=【MakeBOM】反馈&Body=", "", "", 1
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


'库存设置
Private Sub Combo1_LostFocus()
    Select Case Combo1.Text
    Case "TP1"

    Case "TP2"

    Case "TP3"

    Case Else
        MsgBox "不支持的库存类型！请重新选择", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
        Combo1.Text = "TP1"
    End Select

End Sub

'全选
Private Sub Label1_Click()
    Dim PerBox As Object
    For Each PerBox In Me.Controls
        If TypeOf PerBox Is CheckBox Then
            PerBox.Value = Checked
        End If
    DoEvents: Next
End Sub

'反选
Private Sub Label2_Click()
    Dim PerBox As Object
    For Each PerBox In Me.Controls
        If TypeOf PerBox Is CheckBox Then
            If PerBox.Value = Checked Then
                PerBox.Value = Unchecked
            Else
                PerBox.Value = Checked
            End If
        End If
    DoEvents: Next
End Sub

'清空
Private Sub Label3_Click()
    Dim PerBox As Object
    For Each PerBox In Me.Controls
        If TypeOf PerBox Is CheckBox Then
            PerBox.Value = Unchecked
        End If
    DoEvents: Next
End Sub

'手型鼠标实现
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static hCursor As Long
    If hCursor = 0 Then hCursor = LoadCursorBynum&(0&, IDC_HAND)
    SetCursor hCursor
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static hCursor As Long
    If hCursor = 0 Then hCursor = LoadCursorBynum&(0&, IDC_HAND)
    SetCursor hCursor
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static hCursor As Long
    If hCursor = 0 Then hCursor = LoadCursorBynum&(0&, IDC_HAND)
    SetCursor hCursor
End Sub

'允许拖放操作
Private Sub Command_ImportBom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '允许拖放操作
    Dim PLResultFile As Variant
    Dim filePath As String
    
    On Error Resume Next
    For Each PLResultFile In Data.Files
        filePath = PLResultFile
    Next
    
    If filePath = "" Then
        Exit Sub
    End If
    
    Dim filetype As String
    filetype = Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)
    
    Select Case LCase(filetype)
    Case ".bom":
        BuildProjectPath filePath
        BomStage_One
        
    Case ".tsv":
        tsvFilePath = filePath
        
        If BomFilePath = "" Then
            MsgBox "请先选择BOM文件所在路径！", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
            Exit Sub
        End If
        
        BomStage_One
        
    Case ".xls":
        '导入BOM检查器
        BomChecker filePath
        
    Case Else
        MsgBox "文件类型错误！", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "警告"
        ClearPath
    End Select
    
End Sub

'主控制按钮命令
Private Sub Command_ImportBom_Click()
    Dim GetPath As String
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "请选择BOM文件"
    CommonDialog1.Filter = "All File(*.*)|*.*|BOM 文件(*.BOM; *.xls)|*.BOM;*.xls"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    
    GetPath = CommonDialog1.FileName
    
    If GetPath = "" Then
        Exit Sub
    End If
    
    Dim filetype As String
    filetype = Right(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)

    Select Case LCase(filetype)
    Case ".bom":
        BuildProjectPath GetPath
        BomStage_One
        
    Case ".tsv":
        tsvFilePath = GetPath
        
        If BomFilePath = "" Then
            MsgBox "请先选择.BOM文件所在路径！", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
            Exit Sub
        End If
        
        BomStage_One
        
    Case ".xls":
        '导入BOM检查器
        BomChecker GetPath
        
    Case Else
        MsgBox "文件类型错误！", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "警告"
        ClearPath
    End Select
    
End Sub


'*************************************************************************
'**函 数 名：BomStage_One
'**输    入：无
'**输    出：无
'**功能描述：生成BOM Maker File阶段
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-31 23:48:40
'**修 改 人：
'**日    期：
'**版    本：V3.6.3
'*************************************************************************
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
    Process 5, "创建批量查询文件 ..."
    BomMakePLExcel

    '填充来自orCAD BOM的数据并且创建新的.bmf文件
    Process 8, "创建批量查询文件 ..."
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
            istsv = Right$(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)

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
    Process 75, "自动转换BMF文件格式为ANSI ..."
    BmfToAnsi

    '查找bmf文件中有料号，但是没有物料描述的行
    '给出提示是否自动联网更新物料描述
    '本功能确定可以实现 但未完成
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
'          ：流程如下:
'          ：1.根据模版创建Excel BOM
'          ：2.根据需要调整Excel 格式
'          ：3.读取bmf(BOM Maker File)文件 将信息填入Excel
'          ：4.根据信息调整Excel格式
'          ：5.扫描Excel格式 修正部分格式
'          ：6.完成
'          ：注意：领料BOM中的库存信息必须保证是最新的。
'          ：      因此程序会检查tsv文件的产生时间
'          ：      时间不在三天内会提示，重新查询�
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
        Process 80, "创建预BOM ..."
        ExcelCreate BOM_预
        CreateBOM BOM_预
        
    End If
    
    If CheckNcDbg.Value = Checked Then
        Process 83, "创建NC_DBG BOM ..."
        ExcelCreate BOM_NCDBG
        CreateBOM BOM_NCDBG
        
    End If
    
    If CheckNone.Value = Checked Then
        Process 85, "创建NONE BOM ..."
        ExcelCreate BOM_NONE
        CreateBOM BOM_NONE
        
    End If
    
    If Check_调试.Value = Checked Then
        Process 90, "创建调试BOM ..."
        ExcelCreate BOM_调试
        CreateBOM BOM_调试
        
    End If
    
    If Check_生产.Value = Checked Then
        Process 95, "创建生产BOM ..."
        ExcelCreate BOM_生产
        CreateBOM BOM_生产
        
    End If
    
    If Check_领料.Value = Checked Then
        Process 98, "创建领料BOM ..."
        'tsv文件是否失效？
        If DateDiff("d", CDate(GetFileWriteTime(tsvFilePath)), Now) > 3 Then
            MsgBox "tsv文件已经过期[" & GetFileWriteTime(tsvFilePath) & "]！" & vbCrLf & vbCrLf & _
                   "生成领料BOM需最新的tsv文件！", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "警告"
        Else
            ExcelCreate BOM_领料
            CreateBOM BOM_领料
        End If
        
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
