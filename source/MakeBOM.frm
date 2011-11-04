VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MakeBOM(BOMÉú³É¹¤¾ß)"
   ClientHeight    =   4170
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "ËÎÌå"
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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "BOMÀàĞÍ"
      Height          =   1035
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   4155
      Begin VB.CheckBox CheckNcDbg 
         Caption         =   "NC DBGÔª¼ş"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   300
         Width           =   1395
      End
      Begin VB.CheckBox CheckAll 
         Caption         =   "È«Ñ¡"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1035
      End
      Begin VB.CheckBox Check_Éú²ú 
         Caption         =   "Éú²úBOM"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox Check_µ÷ÊÔ 
         Caption         =   "µ÷ÊÔBOM"
         Height          =   255
         Left            =   1380
         TabIndex        =   9
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox Check_ÁìÁÏ 
         Caption         =   "ÁìÁÏBOM"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   1155
      End
      Begin VB.CheckBox CheckPreBom 
         Caption         =   "Ô¤BOM"
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
      Caption         =   "¿â´æÀàĞÍ"
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
      Caption         =   "»úĞÍÃû³Æ"
      Height          =   615
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.TextBox ItemNameText 
         BeginProperty Font 
            Name            =   "ËÎÌå"
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
            Object.ToolTipText     =   "³ÌĞòÔËĞĞ×´Ì¬ÃèÊö"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1693
            MinWidth        =   1693
            Text            =   "0%"
            TextSave        =   "0%"
            Key             =   "process_text"
            Object.ToolTipText     =   "Ö´ĞĞ½ø¶È"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command_ImportBom 
      Caption         =   "Éú³ÉBOM"
      BeginProperty Font 
         Name            =   "ËÎÌå"
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
      Caption         =   "·â×°¿â"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu menu_checker 
      Caption         =   "BomChecker"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu menu_feedback 
      Caption         =   "·´À¡"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
   End
   Begin VB.Menu menu_about 
      Caption         =   "¹ØÓÚ"
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
'**Ä£ ¿é Ãû£ºMainForm
'**Ëµ    Ã÷£ºTP-LINK SMB Switch Product Line Hardware Group
'**          °æÈ¨ËùÓĞ2011 - 2012(C)
'**
'**´´ ½¨ ÈË£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-22 12:08:02
'**ĞŞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**Ãè    Êö£º´°¿ÚÖ÷ÌåÊµÏÖÎÄ¼ş£¬½ö°üº¬´°¿Ú¹¦ÄÜÊµÏÖ´úÂëÒÔ¼°Ö´ĞĞÁ÷³Ì¿ØÖÆ
'**°æ    ±¾£ºV3.2.38
'*************************************************************************

Option Explicit

'ÔØÈë³ÌĞòÅäÖÃ
Private Sub Form_Load()
    
    Dim X As Long
    Dim Y As Long
    
    '³õÊ¼»¯Êı¾İ¿â
    If InitLib(App.Path & "\STD.lst") = False Then
        Command_ImportBom.Enabled = False
    End If
    
    '»ñÈ¡³ÌĞòÉèÖÃ
    ItemName = GetRegValue(App.EXEName, "Product", "")
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    ItemNameText.Text = ItemName
    Combo1.Text = GetRegValue(App.EXEName, "Storage", "TP1")
    
    '³õÊ¼»¯´°¿ÚÎ»ÖÃ Ä¬ÈÏÔÚÆÁÄ»ÖĞÑë
    X = GetRegValue(App.EXEName, "WinLeft", Screen.Width / 2 - Me.Width / 2)
    Y = GetRegValue(App.EXEName, "WinTop", Screen.Height / 2 - Me.Height / 2)
        
    If X > Screen.Width Or Y > Screen.Height Or _
       X > 0 Or Y > 0 Then
        X = Screen.Width / 2 - Me.Width / 2
        Y = Screen.Height / 2 - Me.Height / 2
    End If
    
    Me.Move X, Y
    
    '»ñÈ¡´°¿Ú×´Ì¬
    If GetRegValue(App.EXEName, "OnTop", 1) = 1 Then
        menu_winpos.Caption = "--"
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
    Command_ImportBom.Caption = "Éú³ÉBOM" & vbCrLf & vbCrLf & "£¨BomChecker£©"
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ö±½Óµ÷ÓÃ°´Å¥µÄÍÏ·ÅĞ§¹û
    Command_ImportBom_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '³ÌĞòÅäÖÃÊı¾İ
    SetRegValue App.EXEName, "Storage", iREG_SZ, Combo1.Text
    
    SetRegValue App.EXEName, "ProjectDir", iREG_SZ, ProjectDir
    SetRegValue App.EXEName, "Product", iREG_SZ, ItemNameText.Text
    
    '´°¿ÚÎ»ÖÃ
    SetRegValue App.EXEName, "WinLeft", iREG_DWORD, Me.Left
    SetRegValue App.EXEName, "WinTop", iREG_DWORD, Me.Top
    
    '´°¿Ú×´Ì¬
    If menu_winpos.Caption = "|" Then
        SetRegValue App.EXEName, "OnTop", iREG_DWORD, 0
    Else
        SetRegValue App.EXEName, "OnTop", iREG_DWORD, 1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '½áÊøËùÓĞµÄ´°Ìå
    Dim counter As Integer
    For counter = 0 To Forms.Count - 1
        Unload Forms(counter)
    Next
    
    End
    
End Sub

'²Ëµ¥
Private Sub menu_lib_Click()
    '´ò¿ª¿âÎÄ¼ş
    frmLib.Show 1
End Sub

Private Sub menu_checker_Click()
    Dim GetPath As String
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "ÇëÑ¡ÔñExcel¸ñÊ½µÄBOMÎÄ¼ş"
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
        'µ¼ÈëBOM¼ì²éÆ÷
        BomChecker GetPath
        
    Case Else
        MsgBox "ÎÄ¼şÀàĞÍ´íÎó£¡", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "¾¯¸æ"
        ClearPath
    End Select
    
End Sub

Private Sub menu_about_Click()
    frmAbout.Show 1
End Sub

Private Sub menu_feedback_Click()
    ShellExecute 0, "open", "mailto:shenhao@tp-link.net?subject=¡¾MakeBOM¡¿·´À¡&Body=", "", "", 1
End Sub

Private Sub menu_winpos_Click()
    'ÉèÖÃ´°¿ÚÊÇ·ñ¹Ì¶¨ÔÚ×îÉÏ²ã
    If menu_winpos.Caption = "|" Then
        menu_winpos.Caption = "--"
        '½« ´°¿ÚÉè¶¨³ÉÓÀÔ¶±£³ÖÔÚ×îÉÏ²ã
        SetWindowsPos_TopMost Me.hwnd
    Else
        menu_winpos.Caption = "|"
        'È¡Ïû×îÉÏ²ãÉè¶¨
        SetWindowsPos_NoTopMost Me.hwnd
    End If
    
End Sub


'¿â´æÉèÖÃ
Private Sub Combo1_LostFocus()
    Select Case Combo1.Text
    Case "TP1"

    Case "TP2"

    Case "TP3"

    Case Else
        MsgBox "²»Ö§³ÖµÄ¿â´æÀàĞÍ£¡ÇëÖØĞÂÑ¡Ôñ", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
        Combo1.Text = "TP1"
    End Select

End Sub

'Ñ¡ÔñĞèÒªÉú³ÉµÄBOMÎÄ¼ş
Private Sub CheckAll_Click()

    If CheckAll.Value = Checked Then
        CheckPreBom.Value = Checked
        CheckNcDbg.Value = Checked
        Check_ÁìÁÏ.Value = Checked
        Check_µ÷ÊÔ.Value = Checked
        Check_Éú²ú.Value = Checked
    End If

End Sub

Private Sub CheckCheck()
    If CheckPreBom.Value = Checked And _
       CheckNcDbg.Value = Checked And _
       Check_ÁìÁÏ.Value = Checked And _
       Check_µ÷ÊÔ.Value = Checked And _
       Check_Éú²ú.Value = Checked Then
       
        CheckAll.Value = Checked
    Else
        CheckAll.Value = Unchecked
    End If
End Sub

Private Sub Check_µ÷ÊÔ_Click()
    CheckCheck
End Sub

Private Sub Check_ÁìÁÏ_Click()
    CheckCheck
End Sub

Private Sub Check_Éú²ú_Click()
    CheckCheck
End Sub

Private Sub CheckNcDbg_Click()
    CheckCheck
End Sub

Private Sub CheckPreBom_Click()
    CheckCheck
End Sub

'ÔÊĞíÍÏ·Å²Ù×÷
Private Sub Command_ImportBom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ÔÊĞíÍÏ·Å²Ù×÷
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
            MsgBox "ÇëÏÈÑ¡ÔñBOMÎÄ¼şËùÔÚÂ·¾¶£¡", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "ÌáÊ¾"
            Exit Sub
        End If
        
        BomStage_One
        
    Case ".xls":
        'µ¼ÈëBOM¼ì²éÆ÷
        BomChecker filePath
        
    Case Else
        MsgBox "ÎÄ¼şÀàĞÍ´íÎó£¡", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "¾¯¸æ"
        ClearPath
    End Select
    
End Sub

'Ö÷¿ØÖÆ°´Å¥ÃüÁî
Private Sub Command_ImportBom_Click()
    Dim GetPath As String
    ProjectDir = GetRegValue(App.EXEName, "ProjectDir", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "ÇëÑ¡ÔñBOMÎÄ¼ş"
    CommonDialog1.Filter = "All File(*.*)|*.*|BOM ÎÄ¼ş(*.BOM; *.xls)|*.BOM;*.xls"
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
            MsgBox "ÇëÏÈÑ¡Ôñ.BOMÎÄ¼şËùÔÚÂ·¾¶£¡", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "ÌáÊ¾"
            Exit Sub
        End If
        
        BomStage_One
        
    Case ".xls":
        'µ¼ÈëBOM¼ì²éÆ÷
        BomChecker GetPath
        
    Case Else
        MsgBox "ÎÄ¼şÀàĞÍ´íÎó£¡", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "¾¯¸æ"
        ClearPath
    End Select
    
End Sub


'*************************************************************************
'**º¯ Êı Ãû£ºBomStage_One
'**Êä    Èë£ºÎŞ
'**Êä    ³ö£ºÎŞ
'**¹¦ÄÜÃèÊö£ºÉú³ÉBOM Maker File½×¶Î
'**È«¾Ö±äÁ¿£º
'**µ÷ÓÃÄ£¿é£º
'**×÷    Õß£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-31 23:48:40
'**ĞŞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**°æ    ±¾£ºV3.6.3
'*************************************************************************
Private Sub BomStage_One()

    '==============================================
    '±¾½×¶Î½«»áÉú³ÉBOM Maker File
    '1.µ¼ÈëtsvÎÄ¼şĞÅÏ¢µ½bmfÎÄ¼şÖĞ
    '2.±£ÁôÔø¾­ÓĞµÄ²¿·ÖĞÅÏ¢ ÈçÃèÊö ¿â´æµÈ
    '3.½«ËùĞèĞÅÏ¢ÕûÀíÎª±ê×¼¸ñÊ½ ±ãÓÚºóÒ»½×¶Î¶ÁÈ¡
    '4.Ê¹ÓÃÎÄ±¾¸ñÊ½±ãÓÚ°æ±¾¿ØÖÆ
    '==============================================

    Dim GetPath As String

    KillBom

    Process 2, "¶ÁÈ¡.BOMÎÄ¼şĞÅÏ¢ ..."
    '¶ÁÈ¡.BOMÎÄ¼şĞÅÏ¢
    If ReadBomFile = False Then
        Process 100, ".BOMÎÄ¼şĞÅÏ¢ ...¶ÁÈ¡´íÎó£¡"
        ClearPath
        Exit Sub
    End If

    '´´½¨ÅúÁ¿²éÑ¯ÎÄ¼ş
    Process 5, "´´½¨ÅúÁ¿²éÑ¯ÎÄ¼ş ..."
    BomMakePLExcel

    'Ìî³äÀ´×ÔorCAD BOMµÄÊı¾İ²¢ÇÒ´´½¨ĞÂµÄ.bmfÎÄ¼ş
    Process 8, "´´½¨ÅúÁ¿²éÑ¯ÎÄ¼ş ..."
    BmfMaker

    'Ä¬ÈÏtsvÎÄ¼şÔÚ¹¤×÷Ä¿Â¼ÏÂ
    tsvFilePath = ProjectDir + "fnd_gfm.tsv"

    '²é¿´.BOMÄ¿Â¼ÏÂÊÇ·ñÓĞtsvÎÄ¼ş£¬ÓĞµÄ»°Ö±½Óµ¼Èë Ã»ÓĞ¾ÍÑ¯ÎÊÊÇ·ñ½øÈëERP²éÑ¯
    If Dir(tsvFilePath) = "" Then

        Dim resultL As VbMsgBoxResult
        resultL = MsgBox("ÔÚ¹¤×÷Ä¿Â¼ÏÂÎ´ÕÒµ½ºÏ·¨µÄÅúÁ¿²éÑ¯½á¹ûÎÄ¼ş£¡" & vbCrLf & vbCrLf & vbCrLf & _
                  "ÊÇ·ñ´ò¿ªERPÏµÍ³½øĞĞÅúÁ¿²éÑ¯£¿" & vbCrLf & vbCrLf & _
                  "ÊÇ-µÇÂ¼ERPÏµÍ³¿ªÊ¼²éÑ¯" & vbCrLf & vbCrLf & _
                  "·ñ-Ñ¡ÔñTSVÎÄ¼şÂ·¾¶" & vbCrLf, _
                  vbQuestion + vbMsgBoxSetForeground + vbYesNoCancel)

        If resultL = vbYes Then

            AutoLoginERP "RD_ENGINEER", "123456"
            'FindERP
            Exit Sub

        ElseIf resultL = vbNo Then

            CommonDialog1.FileName = ""
            CommonDialog1.DialogTitle = "ÇëÑ¡Ôñ.tsvÎÄ¼ş"
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
                MsgBox "±ØĞëÎª.tsvÎÄ¼ş£¡", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "¾¯¸æ"
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

    'µ¼ÈëtsvÎÄ¼şÄÚĞÅÏ¢
    ImportTSV

    '×ª»»BMFÎÄ¼ş¸ñÊ½±ÜÃâ³öÏÖÂÒÂë
    Process 75, "×Ô¶¯×ª»»BMFÎÄ¼ş¸ñÊ½ÎªANSI ..."
    BmfToAnsi

    '²éÕÒbmfÎÄ¼şÖĞÓĞÁÏºÅ£¬µ«ÊÇÃ»ÓĞÎïÁÏÃèÊöµÄĞĞ
    '¸ø³öÌáÊ¾ÊÇ·ñ×Ô¶¯ÁªÍø¸üĞÂÎïÁÏÃèÊö
    '±¾¹¦ÄÜÈ·¶¨¿ÉÒÔÊµÏÖ µ«Î´Íê³É
    'GetInfoFromERP
    'GetInfoFromERP "RD_ENGINEER", "123456"

    'Ö±½Ó½øÈëµÚ2½×¶Î Éú³ÉExcel¸ñÊ½BOM½×¶Î
    BomStage_Two

End Sub


'*************************************************************************
'**º¯ Êı Ãû£ºBomStage_Two
'**Êä    Èë£ºÎŞ
'**Êä    ³ö£ºÎŞ
'**¹¦ÄÜÃèÊö£º¸ù¾İCheckBoxµÄ×´Ì¬´´½¨ExcelÎÄ¼şºÍÉú³ÉÏàÓ¦µÄBOM
'          £ºÁ÷³ÌÈçÏÂ:
'          £º1.¸ù¾İÄ£°æ´´½¨Excel BOM
'          £º2.¸ù¾İĞèÒªµ÷ÕûExcel ¸ñÊ½
'          £º3.¶ÁÈ¡bmf(BOM Maker File)ÎÄ¼ş ½«ĞÅÏ¢ÌîÈëExcel
'          £º4.¸ù¾İĞÅÏ¢µ÷ÕûExcel¸ñÊ½
'          £º5.É¨ÃèExcel¸ñÊ½ ĞŞÕı²¿·Ö¸ñÊ½
'          £º6.Íê³É
'          £º×¢Òâ£ºÁìÁÏBOMÖĞµÄ¿â´æĞÅÏ¢±ØĞë±£Ö¤ÊÇ×îĞÂµÄ¡£
'          £º      Òò´Ë³ÌĞò»á¼ì²étsvÎÄ¼şµÄ²úÉúÊ±¼ä
'          £º      Ê±¼ä²»ÔÚÈıÌìÄÚ»áÌáÊ¾£¬ÖØĞÂ²éÑ¯ÿ
'**È«¾Ö±äÁ¿£º
'**µ÷ÓÃÄ£¿é£º
'**×÷    Õß£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-22 12:11:09
'**ĞŞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**°æ    ±¾£ºV3.2.38
'*************************************************************************
Public Sub BomStage_Two()
    
    'ÓÉÄ£°å´´½¨ExcelÎÄ¼ş ²¢Éú³ÉĞèÒªµÄExcel BOMÎÄ¼ş
    BomCreate
    
     '============================================
    'ÏÔÊ¾½á¹ûĞÅÏ¢ Ôª¼şÊıÁ¿¸öÊı
    '============================================
    'PartNum(0) : NcPartNum
    'PartNum(1) : DbgPartNum
    'PartNum(2) : DbNcPartNum
    
    'PartNum(3) : LeadPartNum
    'PartNum(4) : SmtPartNum
    'PartNum(5) : OtherPartNum
    '============================================
    
    Dim msgstr As String
    msgstr = "             BOM ÎÄ¼ş´´½¨³É¹¦£¡" & vbCrLf & vbCrLf
    msgstr = msgstr + "          ²å×°   Ôª¼ş¸öÊıÎª   £º " & PartNum(3) & vbCrLf
    msgstr = msgstr + "          Ìù×°   Ôª¼ş¸öÊıÎª   £º " & PartNum(4) & vbCrLf
    msgstr = msgstr + "          ÆäËû   Ôª¼ş¸öÊıÎª   £º " & PartNum(5) & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     Ôª¼ş¸öÊıÎª   £º " & PartNum(0) & vbCrLf
    msgstr = msgstr + "          DBG    Ôª¼ş¸öÊıÎª   £º " & PartNum(1) & vbCrLf
    msgstr = msgstr + "          DBG_NC Ôª¼ş¸öÊıÎª   £º " & PartNum(2) & vbCrLf & vbCrLf

    msgstr = msgstr + "    ×¢Òâ£ºÉú³ÉµÄBOMÎÄ¼şĞèÒª¼ì²éĞŞ¸Äºó²Å¿É¹©ÆÀÉó "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "BOM ÎÄ¼ş´´½¨³É¹¦"
    
    '´ò¿ªÉú³ÉµÄBOMÄ¿Â¼
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    'Çå¿Õ³ÌĞòÒÀÀµµÄÂ·¾¶ĞÅÏ¢£¬±ãÓÚÏÂ´Î×ª»»¿ªÊ¼
    ClearPath
    
    Process 100, "Íê³É£¡"
    
    Exit Sub

ErrorHandle:
    
    'É¾³ıÎ´ÄÜ³É¹¦Éú³ÉµÄÎÄ¼ş
    KillBom
    
    ClearPath

End Sub


Private Sub BomCreate()
    
    If CheckPreBom.Value = Checked Then
        Process 80, "´´½¨Ô¤BOM ..."
        ExcelCreate BOM_Ô¤
        CreateBOM BOM_Ô¤
        
    End If
    
    If CheckNcDbg.Value = Checked Then
        Process 83, "´´½¨NC_DBGºÍNONE BOM ..."
        ExcelCreate BOM_NCDBG
        CreateBOM BOM_NCDBG
        
        ExcelCreate BOM_NONE
        CreateBOM BOM_NONE
        
    End If
    
    If Check_ÁìÁÏ.Value = Checked Then
        Process 85, "´´½¨ÁìÁÏBOM ..."
        'tsvÎÄ¼şÊÇ·ñÊ§Ğ§£¿
        If DateDiff("d", CDate(GetFileWriteTime(tsvFilePath)), Now) > 3 Then
            MsgBox "tsvÎÄ¼şÒÑ¾­¹ıÆÚ[" & GetFileWriteTime(tsvFilePath) & "]£¡" & vbCrLf & vbCrLf & _
                   "Éú³ÉÁìÁÏBOMĞè×îĞÂµÄtsvÎÄ¼ş£¡", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "¾¯¸æ"
        Else
            ExcelCreate BOM_ÁìÁÏ
            CreateBOM BOM_ÁìÁÏ
        End If
        
    End If
    
    If Check_µ÷ÊÔ.Value = Checked Then
        Process 90, "´´½¨µ÷ÊÔBOM ..."
        ExcelCreate BOM_µ÷ÊÔ
        CreateBOM BOM_µ÷ÊÔ
        
    End If
    
    If Check_Éú²ú.Value = Checked Then
        Process 95, "´´½¨Éú²úBOM ..."
        ExcelCreate BOM_Éú²ú
        CreateBOM BOM_Éú²ú
        
    End If
End Sub

Private Sub KillBom()

    KillExcel SaveAsPath & "_Ô¤BOM_BMF.xls"
    KillExcel SaveAsPath & "_ÅúÁ¿×ÊÔ´²éÑ¯.xls"
    KillExcel SaveAsPath & "_None_PartRef.xls"
    KillExcel SaveAsPath & "_NC_DBG.xls"
    
    KillExcel SaveAsPath & "_ÁìÁÏBOM.xls"
    KillExcel SaveAsPath & "_Éú²úBOM.xls"
    KillExcel SaveAsPath & "_µ÷ÊÔBOM.xls"
    
End Sub
