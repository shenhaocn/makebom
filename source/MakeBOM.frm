VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MakeBOM(BOM���ɹ���)"
   ClientHeight    =   4170
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4545
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
   ScaleHeight     =   4170
   ScaleWidth      =   4545
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "BOM����"
      Height          =   1035
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   4155
      Begin VB.CheckBox CheckNcDbg 
         Caption         =   "NC DBGԪ��"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   300
         Width           =   1395
      End
      Begin VB.CheckBox CheckAll 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox Check_���� 
         Caption         =   "����BOM"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox Check_���� 
         Caption         =   "����BOM"
         Height          =   255
         Left            =   1380
         TabIndex        =   9
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox Check_���� 
         Caption         =   "����BOM"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox CheckPreBom 
         Caption         =   "ԤBOM"
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
      Caption         =   "�������"
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
      Caption         =   "��������"
      Height          =   615
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   2775
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
      Top             =   3855
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
            Object.ToolTipText     =   "��������״̬����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1693
            MinWidth        =   1693
            Text            =   "0%"
            TextSave        =   "0%"
            Key             =   "process_text"
            Object.ToolTipText     =   "ִ�н���"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command_ImportBom 
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
      Height          =   1455
      Left            =   180
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   2100
      Width           =   4155
   End
   Begin VB.Menu menu_lib 
      Caption         =   "��װ��"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu menu_checker 
      Caption         =   "BomChecker"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu menu_feedback 
      Caption         =   "����"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
   End
   Begin VB.Menu menu_about 
      Caption         =   "����"
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
'**ģ �� ����MainForm
'**˵    ����TP-LINK SMB Switch Product Line Hardware Group
'**          ��Ȩ����2011 - 2012(C)
'**
'**�� �� �ˣ�Shenhao
'**��    �ڣ�2011-10-22 12:08:02
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ������������ʵ���ļ������������ڹ���ʵ�ִ����Լ�ִ�����̿���
'**��    ����V3.2.38
'*************************************************************************

Option Explicit

'�����������
Private Sub Form_Load()
    
    Dim X As Long
    Dim Y As Long
    
    '��ʼ�����ݿ�
    If InitLib(App.Path & "\STD.lst") = False Then
        Command_ImportBom.Enabled = False
    End If
    
    '��ȡ��������
    ItemName = GetRegValue(App.EXEName, "Product", "")
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    ItemNameText.Text = ItemName
    Combo1.Text = GetRegValue(App.EXEName, "Storage", "TP1")
    
    '��ʼ������λ�� Ĭ������Ļ����
    X = GetRegValue(App.EXEName, "WinLeft", Screen.Width / 2 - Me.Width / 2)
    Y = GetRegValue(App.EXEName, "WinTop", Screen.Height / 2 - Me.Height / 2)
        
    If X > Screen.Width Or Y > Screen.Height Or _
       X > 0 Or Y > 0 Then
        X = Screen.Width / 2 - Me.Width / 2
        Y = Screen.Height / 2 - Me.Height / 2
    End If
    
    Me.Move X, Y
    
    '��ȡ����״̬
    If GetRegValue(App.EXEName, "OnTop", 1) = 1 Then
        menu_winpos.Caption = "--"
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
    Command_ImportBom.Caption = "����BOM" & vbCrLf & vbCrLf & "��BomChecker��"
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ֱ�ӵ��ð�ť���Ϸ�Ч��
    Command_ImportBom_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '������������
    SetRegValue App.EXEName, "Storage", iREG_SZ, Combo1.Text
    
    SetRegValue App.EXEName, "ProjectDir", iREG_SZ, ProjectDir
    SetRegValue App.EXEName, "Product", iREG_SZ, ItemNameText.Text
    
    '����λ��
    SetRegValue App.EXEName, "WinLeft", iREG_DWORD, Me.Left
    SetRegValue App.EXEName, "WinTop", iREG_DWORD, Me.Top
    
    '����״̬
    If menu_winpos.Caption = "|" Then
        SetRegValue App.EXEName, "OnTop", iREG_DWORD, 0
    Else
        SetRegValue App.EXEName, "OnTop", iREG_DWORD, 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '�������еĴ���
    Dim counter As Integer
    For counter = 0 To Forms.Count - 1
        Unload Forms(counter)
    Next
    
    End
    
End Sub

'�˵�
Private Sub menu_lib_Click()
    '�򿪿��ļ�
    frmLib.Show 1
End Sub

Private Sub menu_checker_Click()
    Dim GetPath As String
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "��ѡ��Excel��ʽ��BOM�ļ�"
    CommonDialog1.Filter = "All File(*.*)|*.*|Excel BOM files(*.xls)|*.xls"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    
    GetPath = CommonDialog1.FileName
    
    If GetPath = "" Then
        Exit Sub
    End If
    
    Dim filetype As String
    filetype = Right(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)

    Select Case LCase(filetype)
    Case ".xls":
        '����BOM�����
        BomChecker GetPath
        
    Case Else
        MsgBox "�ļ����ʹ���", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "����"
        ClearPath
    End Select
    
End Sub

Private Sub menu_about_Click()
    frmAbout.Show 1
End Sub

Private Sub menu_feedback_Click()
    ShellExecute 0, "open", "mailto:shenhao@tp-link.net?subject=��MakeBOM������&Body=", "", "", 1
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


'�������
Private Sub Combo1_LostFocus()
    Select Case Combo1.Text
    Case "TP1"

    Case "TP2"

    Case "TP3"

    Case Else
        MsgBox "��֧�ֵĿ�����ͣ�������ѡ��", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
        Combo1.Text = "TP1"
    End Select

End Sub

'ѡ����Ҫ���ɵ�BOM�ļ�
Private Sub CheckAll_Click()

    If CheckAll.Value = Checked Then
        CheckPreBom.Value = Checked
        CheckNcDbg.Value = Checked
        Check_����.Value = Checked
        Check_����.Value = Checked
        Check_����.Value = Checked
    End If

End Sub

Private Sub CheckCheck()
    If CheckPreBom.Value = Checked And _
       CheckNcDbg.Value = Checked And _
       Check_����.Value = Checked And _
       Check_����.Value = Checked And _
       Check_����.Value = Checked Then
       
        CheckAll.Value = Checked
    Else
        CheckAll.Value = Unchecked
    End If
End Sub

Private Sub Check_����_Click()
    CheckCheck
End Sub

Private Sub Check_����_Click()
    CheckCheck
End Sub

Private Sub Check_����_Click()
    CheckCheck
End Sub

Private Sub CheckNcDbg_Click()
    CheckCheck
End Sub

Private Sub CheckPreBom_Click()
    CheckCheck
End Sub

'�����ϷŲ���
Private Sub Command_ImportBom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�����ϷŲ���
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
            MsgBox "����ѡ��BOM�ļ�����·����", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
            Exit Sub
        End If
        
        BomStage_One
        
    Case ".xls":
        '����BOM�����
        BomChecker filePath
        
    Case Else
        MsgBox "�ļ����ʹ���", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "����"
        ClearPath
    End Select
    
End Sub

'�����ư�ť����
Private Sub Command_ImportBom_Click()
    Dim GetPath As String
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "��ѡ��BOM�ļ�"
    CommonDialog1.Filter = "All File(*.*)|*.*|BOM �ļ�(*.BOM; *.xls)|*.BOM;*.xls"
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
            MsgBox "����ѡ��.BOM�ļ�����·����", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
            Exit Sub
        End If
        
        BomStage_One
        
    Case ".xls":
        '����BOM�����
        BomChecker GetPath
        
    Case Else
        MsgBox "�ļ����ʹ���", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "����"
        ClearPath
    End Select
    
End Sub


'*************************************************************************
'**�� �� ����BomStage_One
'**��    �룺��
'**��    ������
'**��������������BOM Maker File�׶�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-31 23:48:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.6.3
'*************************************************************************
Private Sub BomStage_One()

    '==============================================
    '���׶ν�������BOM Maker File
    '1.����tsv�ļ���Ϣ��bmf�ļ���
    '2.���������еĲ�����Ϣ ������ ����
    '3.��������Ϣ����Ϊ��׼��ʽ ���ں�һ�׶ζ�ȡ
    '4.ʹ���ı���ʽ���ڰ汾����
    '==============================================

    Dim GetPath As String

    KillBom

    Process 2, "��ȡ.BOM�ļ���Ϣ ..."
    '��ȡ.BOM�ļ���Ϣ
    If ReadBomFile = False Then
        Process 100, ".BOM�ļ���Ϣ ...��ȡ����"
        ClearPath
        Exit Sub
    End If

    '����������ѯ�ļ�
    Process 5, "����������ѯ�ļ� ..."
    BomMakePLExcel

    '�������orCAD BOM�����ݲ��Ҵ����µ�.bmf�ļ�
    Process 8, "����������ѯ�ļ� ..."
    BmfMaker

    'Ĭ��tsv�ļ��ڹ���Ŀ¼��
    tsvFilePath = ProjectDir + "fnd_gfm.tsv"

    '�鿴.BOMĿ¼���Ƿ���tsv�ļ����еĻ�ֱ�ӵ��� û�о�ѯ���Ƿ����ERP��ѯ
    If Dir(tsvFilePath) = "" Then

        Dim resultL As VbMsgBoxResult
        resultL = MsgBox("�ڹ���Ŀ¼��δ�ҵ��Ϸ���������ѯ����ļ���" & vbCrLf & vbCrLf & vbCrLf & _
                  "�Ƿ��ERPϵͳ����������ѯ��" & vbCrLf & vbCrLf & _
                  "��-��¼ERPϵͳ��ʼ��ѯ" & vbCrLf & vbCrLf & _
                  "��-ѡ��TSV�ļ�·��" & vbCrLf, _
                  vbQuestion + vbMsgBoxSetForeground + vbYesNoCancel)

        If resultL = vbYes Then

            AutoLoginERP "RD_ENGINEER", "123456"
            'FindERP
            Exit Sub

        ElseIf resultL = vbNo Then

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
            istsv = Right$(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)

            If istsv <> ".tsv" Then
                MsgBox "����Ϊ.tsv�ļ���", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "����"
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

    '����tsv�ļ�����Ϣ
    ImportTSV

    'ת��BMF�ļ���ʽ�����������
    Process 75, "�Զ�ת��BMF�ļ���ʽΪANSI ..."
    BmfToAnsi

    '����bmf�ļ������Ϻţ�����û��������������
    '������ʾ�Ƿ��Զ�����������������
    '������ȷ������ʵ�� ��δ���
    'GetInfoFromERP
    'GetInfoFromERP "RD_ENGINEER", "123456"

    'ֱ�ӽ����2�׶� ����Excel��ʽBOM�׶�
    BomStage_Two

End Sub


'*************************************************************************
'**�� �� ����BomStage_Two
'**��    �룺��
'**��    ������
'**��������������CheckBox��״̬����Excel�ļ���������Ӧ��BOM
'          ����������:
'          ��1.����ģ�洴��Excel BOM
'          ��2.������Ҫ����Excel ��ʽ
'          ��3.��ȡbmf(BOM Maker File)�ļ� ����Ϣ����Excel
'          ��4.������Ϣ����Excel��ʽ
'          ��5.ɨ��Excel��ʽ �������ָ�ʽ
'          ��6.���
'          ��ע�⣺����BOM�еĿ����Ϣ���뱣֤�����µġ�
'          ��      ��˳������tsv�ļ��Ĳ���ʱ��
'          ��      ʱ�䲻�������ڻ���ʾ�����²�ѯ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Shenhao
'**��    �ڣ�2011-10-22 12:11:09
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V3.2.38
'*************************************************************************
Public Sub BomStage_Two()
    
    '��ģ�崴��Excel�ļ� ��������Ҫ��Excel BOM�ļ�
    BomCreate
    
     '============================================
    '��ʾ�����Ϣ Ԫ����������
    '============================================
    'PartNum(0) : NcPartNum
    'PartNum(1) : DbgPartNum
    'PartNum(2) : DbNcPartNum
    
    'PartNum(3) : LeadPartNum
    'PartNum(4) : SmtPartNum
    'PartNum(5) : OtherPartNum
    '============================================
    
    Dim msgstr As String
    msgstr = "             BOM �ļ������ɹ���" & vbCrLf & vbCrLf
    msgstr = msgstr + "          ��װ   Ԫ������Ϊ   �� " & PartNum(3) & vbCrLf
    msgstr = msgstr + "          ��װ   Ԫ������Ϊ   �� " & PartNum(4) & vbCrLf
    msgstr = msgstr + "          ����   Ԫ������Ϊ   �� " & PartNum(5) & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     Ԫ������Ϊ   �� " & PartNum(0) & vbCrLf
    msgstr = msgstr + "          DBG    Ԫ������Ϊ   �� " & PartNum(1) & vbCrLf
    msgstr = msgstr + "          DBG_NC Ԫ������Ϊ   �� " & PartNum(2) & vbCrLf & vbCrLf

    msgstr = msgstr + "    ע�⣺���ɵ�BOM�ļ���Ҫ����޸ĺ�ſɹ����� "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "BOM �ļ������ɹ�"
    
    '�����ɵ�BOMĿ¼
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    '��ճ���������·����Ϣ�������´�ת����ʼ
    ClearPath
    
    Process 100, "��ɣ�"
    
    Exit Sub

ErrorHandle:
    
    'ɾ��δ�ܳɹ����ɵ��ļ�
    KillBom
    
    ClearPath

End Sub


Private Sub BomCreate()
    
    If CheckPreBom.Value = Checked Then
        Process 80, "����ԤBOM ..."
        ExcelCreate BOM_Ԥ
        CreateBOM BOM_Ԥ
        
    End If
    
    If CheckNcDbg.Value = Checked Then
        Process 83, "����NC_DBG��NONE BOM ..."
        ExcelCreate BOM_NCDBG
        CreateBOM BOM_NCDBG
        
        ExcelCreate BOM_NONE
        CreateBOM BOM_NONE
        
    End If
    
    If Check_����.Value = Checked Then
        Process 85, "��������BOM ..."
        'tsv�ļ��Ƿ�ʧЧ��
        If DateDiff("d", CDate(GetFileWriteTime(tsvFilePath)), Now) > 3 Then
            MsgBox "tsv�ļ��Ѿ�����[" & GetFileWriteTime(tsvFilePath) & "]��" & vbCrLf & vbCrLf & _
                   "��������BOM�����µ�tsv�ļ���", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "����"
        Else
            ExcelCreate BOM_����
            CreateBOM BOM_����
        End If
        
    End If
    
    If Check_����.Value = Checked Then
        Process 90, "��������BOM ..."
        ExcelCreate BOM_����
        CreateBOM BOM_����
        
    End If
    
    If Check_����.Value = Checked Then
        Process 95, "��������BOM ..."
        ExcelCreate BOM_����
        CreateBOM BOM_����
        
    End If
End Sub

Private Sub KillBom()

    KillExcel SaveAsPath & "_ԤBOM_BMF.xls"
    KillExcel SaveAsPath & "_������Դ��ѯ.xls"
    KillExcel SaveAsPath & "_None_PartRef.xls"
    KillExcel SaveAsPath & "_NC_DBG.xls"
    
    KillExcel SaveAsPath & "_����BOM.xls"
    KillExcel SaveAsPath & "_����BOM.xls"
    KillExcel SaveAsPath & "_����BOM.xls"
    
End Sub
