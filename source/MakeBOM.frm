VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BOM���ɹ���"
   ClientHeight    =   2070
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      Caption         =   "�������"
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
      Caption         =   "��������"
      Height          =   615
      Left            =   300
      TabIndex        =   4
      Top             =   120
      Width           =   2475
      Begin VB.TextBox ItemNameText 
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "tsv�ļ�����"
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
            Text            =   "��ѡ��.BOM�ļ�..."
            TextSave        =   "��ѡ��.BOM�ļ�..."
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
      Caption         =   "����BOM"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����װ��"
   End
   Begin VB.Menu menu_update 
      Caption         =   "��װ�����"
   End
   Begin VB.Menu menu_feedback 
      Caption         =   "����"
   End
   Begin VB.Menu menu_about 
      Caption         =   "����"
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
'��������ʵ���ļ������������ڹ���ʵ�ִ����Լ�ִ�����̿���
'=========================================================

Private Sub Form_Load()
    
    '��ʼ�����ݿ�
    If InitLib(App.Path & "\STD.lst") = False Then
        MakeBOM_Command.Enabled = False
    End If
    
    '��ȡ�ϴι���Ŀ¼
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "�ϴι���Ŀ¼", "E:\")
    ItemName = GetSetting(App.EXEName, "TaskName", "��һ����Ŀ��������", "")
    
    ItemNameText.Text = ItemName
    
    '��ȡ��������
    UTFCombo.Text = GetSetting(App.EXEName, "tsvEncoder", "tsv�ļ�����", "UTF-8")
    Combo1.Text = GetSetting(App.EXEName, "SelectStorage", "�������", "TP1")
    
    '��ʼ������λ�ú�״̬
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
    'ֱ�ӵ��ð�ť���Ϸ�Ч��
    MakeBOM_Command_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '������������
    SaveSetting App.EXEName, "tsvEncoder", "tsv�ļ�����", UTFCombo.Text
    SaveSetting App.EXEName, "SelectStorage", "�������", Combo1.Text
    
    SaveSetting App.EXEName, "ProjectDir", "�ϴι���Ŀ¼", ProjectDir
    SaveSetting App.EXEName, "TaskName", "��һ����Ŀ��������", ItemNameText.Text
    
    '����λ��
    SaveSetting App.EXEName, "WindowPosition", "Left", Me.Left
    SaveSetting App.EXEName, "WindowPosition", "Top", Me.Top
    SaveSetting App.EXEName, "WindowPosition", "AlwaysOnTop", menu_winpos.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '�������еĴ���
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
    ShellExecute 0, "open", "mailto:shenhao@tp-link.net?subject=��MakeBOM������&Body=", "", "", 1
End Sub

Private Sub menu_lib_Click()
    '�򿪿��ļ�
    Shell "notepad " & LibFilePath, vbMaximizedFocus
End Sub

Private Sub menu_update_Click()
    frmUpdateLib.Show 1
End Sub

Private Sub menu_winpos_Click()
    '���ô����Ƿ�̶������ϲ�
    If menu_winpos.Caption = "|" Then
        menu_winpos.Caption = "--"
        '�� �����趨����Զ���������ϲ�
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        'ȡ�����ϲ��趨
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
End Sub

Private Sub UTFCombo_Click()
    SaveSetting App.EXEName, "tsvEncoder", "tsv�ļ�����", UTFCombo.Text
End Sub

Private Sub UTFCombo_LostFocus()
    Select Case UTFCombo.Text
    Case "UTF-8"
            
    Case "ANSI"
        
    Case "UTF-16LE"
        
    Case "UTF-16BE"
        
    Case Else
        MsgBox "��֧�ֵ��ı��ļ����룡", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
        UTFCombo.Text = GetSetting(App.EXEName, "tsvEncoder", "tsv�ļ�����", "UTF-8")
    End Select
    
    SaveSetting App.EXEName, "tsvEncoder", "tsv�ļ�����", UTFCombo.Text
End Sub

Private Sub Combo1_Click()
    SaveSetting App.EXEName, "SelectStorage", "�������", Combo1.Text
End Sub

Private Sub Combo1_LostFocus()
    Select Case Combo1.Text
    Case "TP1"

    Case "TP2"

    Case "TP3"

    Case Else
        MsgBox "��֧�ֵĿ�����ͣ�������ѡ��", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
        Combo1.Text = "TP1"
    End Select
    
    SaveSetting App.EXEName, "SelectStorage", "�������", Combo1.Text
End Sub


Private Sub MakeBOM_Command_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�����ϷŲ���
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
            MsgBox "����ѡ��.BOM�ļ�����·����", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
            Exit Sub
        End If
        
        BomStage_Two
    End If
End Sub

Private Sub MakeBOM_Command_Click()
    Dim GetPath As String
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "�ϴι���Ŀ¼", "E:\")
    
    If BomFilePath = "" Then
        CommonDialog1.InitDir = ProjectDir
        CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "��ѡ��.BOM�ļ�"
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
            MsgBox "����Ϊ.BOM�ļ���", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "����"
            ClearPath
            Exit Sub
        Else
            BuildProjectPath GetPath
            
            BomStage_One
        End If
    Else
        If tsvFilePath = "" Then
            CommonDialog1.FileName = ""
            CommonDialog1.DialogTitle = "��ѡ��.tsv�ļ�"
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
                MsgBox "����Ϊ.tsv�ļ���", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "����"
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
    
    'ɾ����ʱ���ļ�
    msgstr = SaveAsPath & "_PCBA_BOM.xls" & vbCrLf
    msgstr = msgstr + SaveAsPath & "_������Դ��ѯ.xls" & vbCrLf
    msgstr = msgstr + SaveAsPath & "_None_PartRef.xls" & vbCrLf
    msgstr = msgstr + SaveAsPath & "_NC_DBG.xls" & vbCrLf & vbCrLf
    msgstr = msgstr + "�Ѿ����ڣ�ʹ����Щ�ļ���" & vbCrLf
    
    If Dir(SaveAsPath & "_PCBA_BOM.xls") <> "" _
            Or Dir(SaveAsPath & "_������Դ��ѯ.xls") <> "" _
                Or Dir(SaveAsPath & "_None_PartRef.xls") <> "" _
                Or Dir(SaveAsPath & "_NC_DBG.xls") <> "" Then
        If MsgBox(msgstr, vbInformation + vbMsgBoxSetForeground + vbYesNo) = vbYes Then
            Process 10, "��ȡԪ����..."
            CalcPartNum
            
            Process 20, "��ʼ����BOM��2�׶Σ�ѡ��tsv�ļ�·��..."
            If tsvFilePath = "" Then
                MakeBOM_Command_Click
            Else
                BomStage_Two
            End If
            
            Exit Sub
        Else
            KillExcel SaveAsPath & "_PCBA_BOM.xls"
            KillExcel SaveAsPath & "_������Դ��ѯ.xls"
            KillExcel SaveAsPath & "_None_PartRef.xls"
            KillExcel SaveAsPath & "_NC_DBG.xls"
        End If
    End If
    
    Process 2, "��ȡ.BOM�ļ���Ϣ ..."
    '��ȡ.BOM�ļ���Ϣ
    If ReadBomFile = False Then
        Process 100, ".BOM�ļ���Ϣ ...��ȡ����"
        ClearPath
        Exit Sub
    End If
    
    '��ģ�崴��Excel�ļ�
    ExcelCreate
    
    '�������orCAD BOM������
    BomDraft
    
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    If MsgBox("�Ƿ��ERPϵͳ����������ѯ��", vbQuestion + vbMsgBoxSetForeground + vbYesNo) = vbYes Then
        AutoLoginERP "RD_ENGINEER", "123456"
        'FindERP
    End If
    
End Sub


Public Sub BomStage_Two()
    
    Dim msgstr As String
    'ɾ����ʱ���ļ�
    'msgstr = SaveAsPath & "_����BOM.xls" & vbCrLf
    'msgstr = msgstr + SaveAsPath & "_����BOM.xls" & vbCrLf
    'msgstr = msgstr + SaveAsPath & "_����BOM.xls" & vbCrLf & vbCrLf
    'msgstr = msgstr + "�Ѿ����ڣ��Ƿ�ɾ���Ա��������ɣ�" & vbCrLf
    
    If Dir(SaveAsPath & "_����BOM.xls") <> "" _
            Or Dir(SaveAsPath & "_����BOM.xls") <> "" _
                Or Dir(SaveAsPath & "_����BOM.xls") <> "" Then
        'If MsgBox(msgstr, vbInformation + vbMsgBoxSetForeground + vbYesNo) = vbYes Then
        KillExcel SaveAsPath & "_����BOM.xls"
        KillExcel SaveAsPath & "_����BOM.xls"
        KillExcel SaveAsPath & "_����BOM.xls"
        'Else
        '    Exit Sub
        'End If
    End If
    
    '�������BOM
    If CreateAllBOM = False Then
        Process 100, "����BOM�ļ�ʧ�ܣ�"
        GoTo ErrorHandle
    End If
    '��������BOM��ʽ������
    If BomAdjust = False Then
        Process 100, "����BOM�ļ�ʧ�ܣ�"
        GoTo ErrorHandle
    End If
    
    '�������BOM
    If ImportTSV(SaveAsPath & "_PCBA_BOM.xls", 80) = False Then
        Process 100, "tsv�ļ����������tsv�ļ���"
        GoTo ErrorHandle
    End If
    
    If ImportTSV(SaveAsPath & "_����BOM.xls", 84) = False Then
        Process 100, "tsv�ļ����������tsv�ļ���"
        GoTo ErrorHandle
    End If
    
    If ImportTSV(SaveAsPath & "_����BOM.xls", 88) = False Then
        Process 100, "tsv�ļ����������tsv�ļ���"
        GoTo ErrorHandle
    End If
    
    If ImportTSV(SaveAsPath & "_����BOM.xls", 92) = False Then
        Process 100, "tsv�ļ����������tsv�ļ���"
        GoTo ErrorHandle
    End If
    
    '�����ɵ�BOM���Ա����
    'ShellExecute 0, "open", SaveAsPath & "_PCBA_BOM.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_NC_DBG.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_None_PartRef.xls", "", "", 1
    
    'ShellExecute 0, "open", SaveAsPath & "_����BOM.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_����BOM.xls", "", "", 1
    'ShellExecute 0, "open", SaveAsPath & "_����BOM.xls", "", "", 1
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    '��ճ���������·����Ϣ�������´�ת����ʼ
    ClearPath
    
    Process 100, "��ɣ�"
    Exit Sub

ErrorHandle:
    
    'ɾ��δ�ܳɹ����ɵ��ļ�
    KillExcel SaveAsPath & "_����BOM.xls"
    KillExcel SaveAsPath & "_����BOM.xls"
    KillExcel SaveAsPath & "_����BOM.xls"
    
    ClearPath

End Sub

