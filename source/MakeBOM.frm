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
      Caption         =   "Éú³ÉBOM"
      Height          =   1035
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   4155
      Begin VB.CheckBox CheckNcDbg 
         Caption         =   "NC DBGÔª¼þ"
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
      Caption         =   "¿â´æÀàÐÍ"
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
      Caption         =   "»úÐÍÃû³Æ"
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
            Object.ToolTipText     =   "³ÌÐòÔËÐÐ×´Ì¬ÃèÊö"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1693
            MinWidth        =   1693
            Text            =   "0%"
            TextSave        =   "0%"
            Key             =   "process_text"
            Object.ToolTipText     =   "Ö´ÐÐ½ø¶È"
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
'**          °æÈ¨ËùÓÐ2011 - 2012(C)
'**
'**´´ ½¨ ÈË£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-22 12:08:02
'**ÐÞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**Ãè    Êö£º´°¿ÚÖ÷ÌåÊµÏÖÎÄ¼þ£¬½ö°üº¬´°¿Ú¹¦ÄÜÊµÏÖ´úÂëÒÔ¼°Ö´ÐÐÁ÷³Ì¿ØÖÆ
'**°æ    ±¾£ºV3.2.38
'*************************************************************************

Option Explicit

Private Sub Form_Load()
    
    '³õÊ¼»¯Êý¾Ý¿â
    If InitLib(App.Path & "\STD.lst") = False Then
        Command_ImportBom.Enabled = False
    End If
    
    '»ñÈ¡ÉÏ´Î¹¤×÷Ä¿Â¼
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "ÉÏ´Î¹¤×÷Ä¿Â¼", "E:\")
    ItemName = GetSetting(App.EXEName, "TaskName", "ÉÏÒ»´ÎÏîÄ¿»úÐÍÃû³Æ", "")
    
    ItemNameText.Text = ItemName
    
    '»ñÈ¡³ÌÐòÉèÖÃ
    Combo1.Text = GetSetting(App.EXEName, "SelectStorage", "¿â´æÀàÐÍ", "TP1")
    
    '³õÊ¼»¯´°¿ÚÎ»ÖÃºÍ×´Ì¬
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
    
    Command_ImportBom.Caption = "Éú³ÉBOM" & vbCrLf & vbCrLf & "£¨BomChecker£©"
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Ö±½Óµ÷ÓÃ°´Å¥µÄÍÏ·ÅÐ§¹û
    Command_ImportBom_OLEDragDrop Data, Effect, Button, Shift, X, Y
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '³ÌÐòÅäÖÃÊý¾Ý
    SaveSetting App.EXEName, "SelectStorage", "¿â´æÀàÐÍ", Combo1.Text
    
    SaveSetting App.EXEName, "ProjectDir", "ÉÏ´Î¹¤×÷Ä¿Â¼", ProjectDir
    SaveSetting App.EXEName, "TaskName", "ÉÏÒ»´ÎÏîÄ¿»úÐÍÃû³Æ", ItemNameText.Text
    
    '´°¿ÚÎ»ÖÃ
    SaveSetting App.EXEName, "WindowPosition", "Left", Me.Left
    SaveSetting App.EXEName, "WindowPosition", "Top", Me.Top
    SaveSetting App.EXEName, "WindowPosition", "AlwaysOnTop", menu_winpos.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '½áÊøËùÓÐµÄ´°Ìå
    Dim counter As Integer
    For counter = 0 To Forms.Count - 1
        Unload Forms(counter)
    Next
    
    End
    
End Sub

Private Sub menu_lib_Click()
    '´ò¿ª¿âÎÄ¼þ
    frmLib.Show 1
End Sub

Private Sub menu_checker_Click()
    Dim GetPath As String
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "ÉÏ´Î¹¤×÷Ä¿Â¼", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "ÇëÑ¡ÔñExcel¸ñÊ½µÄBOMÎÄ¼þ"
    CommonDialog1.Filter = "All File(*.*)|*.*|Excel BOM files(*.xls)|*.xls"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.ShowOpen
    
    GetPath = CommonDialog1.FileName
    
    If GetPath = "" Then
        Exit Sub
    End If
    
    Dim isbom As String
    isbom = Right(GetPath, Len(GetPath) - InStrRev(GetPath, ".") + 1)

    If isbom <> ".XLS" And isbom <> ".xls" Then
        MsgBox "±ØÐëÎª.xlsÎÄ¼þ£¡", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "¾¯¸æ"
        Exit Sub
    End If
    
    'µ¼ÈëBOM¼ì²éÆ÷
    BomChecker GetPath
    
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



Private Sub Combo1_Click()
    SaveSetting App.EXEName, "SelectStorage", "¿â´æÀàÐÍ", Combo1.Text
End Sub

Private Sub Combo1_LostFocus()
    Select Case Combo1.Text
    Case "TP1"

    Case "TP2"

    Case "TP3"

    Case Else
        MsgBox "²»Ö§³ÖµÄ¿â´æÀàÐÍ£¡ÇëÖØÐÂÑ¡Ôñ", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
        Combo1.Text = "TP1"
    End Select
    
    SaveSetting App.EXEName, "SelectStorage", "¿â´æÀàÐÍ", Combo1.Text
End Sub

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

Private Sub Command_ImportBom_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ÔÊÐíÍÏ·Å²Ù×÷
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
            MsgBox "ÇëÏÈÑ¡Ôñ.BOMÎÄ¼þËùÔÚÂ·¾¶£¡", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "ÌáÊ¾"
            Exit Sub
        End If
        
        BomStage_One
    End If
    
    If filetype = ".xls" Then
        'µ¼ÈëBOM¼ì²éÆ÷
        BomChecker filePath
    End If
    
End Sub

Private Sub Command_ImportBom_Click()
    Dim GetPath As String
    ProjectDir = GetSetting(App.EXEName, "ProjectDir", "ÉÏ´Î¹¤×÷Ä¿Â¼", "E:\")
    
    CommonDialog1.InitDir = ProjectDir
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "ÇëÑ¡Ôñ.BOMÎÄ¼þ"
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
        MsgBox "±ØÐëÎª.BOMÎÄ¼þ£¡", vbMsgBoxSetForeground + vbExclamation + vbOKOnly, "¾¯¸æ"
        ClearPath
        Exit Sub
    End If
    
    BuildProjectPath GetPath

    BomStage_One
    
End Sub


'*************************************************************************
'**º¯ Êý Ãû£ºBomStage_One
'**Êä    Èë£ºÎÞ
'**Êä    ³ö£ºÎÞ
'**¹¦ÄÜÃèÊö£ºÉú³ÉBOM Maker File½×¶Î
'**È«¾Ö±äÁ¿£º
'**µ÷ÓÃÄ£¿é£º
'**×÷    Õß£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-31 23:48:40
'**ÐÞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**°æ    ±¾£ºV3.6.3
'*************************************************************************
Private Sub BomStage_One()

    '==============================================
    '±¾½×¶Î½«»áÉú³ÉBOM Maker File
    '1.µ¼ÈëtsvÎÄ¼þÐÅÏ¢µ½bmfÎÄ¼þÖÐ
    '2.±£ÁôÔø¾­ÓÐµÄ²¿·ÖÐÅÏ¢ ÈçÃèÊö ¿â´æµÈ
    '3.½«ËùÐèÐÅÏ¢ÕûÀíÎª±ê×¼¸ñÊ½ ±ãÓÚºóÒ»½×¶Î¶ÁÈ¡
    '4.Ê¹ÓÃÎÄ±¾¸ñÊ½±ãÓÚ°æ±¾¿ØÖÆ
    '==============================================

    Dim GetPath As String

    KillBom

    Process 2, "¶ÁÈ¡.BOMÎÄ¼þÐÅÏ¢ ..."
    '¶ÁÈ¡.BOMÎÄ¼þÐÅÏ¢
    If ReadBomFile = False Then
        Process 100, ".BOMÎÄ¼þÐÅÏ¢ ...¶ÁÈ¡´íÎó£¡"
        ClearPath
        Exit Sub
    End If

    '´´½¨ÅúÁ¿²éÑ¯ÎÄ¼þ
    Process 5, "´´½¨ÅúÁ¿²éÑ¯ÎÄ¼þ ..."
    BomMakePLExcel

    'Ìî³äÀ´×ÔorCAD BOMµÄÊý¾Ý²¢ÇÒ´´½¨ÐÂµÄ.bmfÎÄ¼þ
    Process 8, "´´½¨ÅúÁ¿²éÑ¯ÎÄ¼þ ..."
    BmfMaker

    'Ä¬ÈÏtsvÎÄ¼þÔÚ¹¤×÷Ä¿Â¼ÏÂ
    tsvFilePath = ProjectDir + "fnd_gfm.tsv"

    '²é¿´.BOMÄ¿Â¼ÏÂÊÇ·ñÓÐtsvÎÄ¼þ£¬ÓÐµÄ»°Ö±½Óµ¼Èë Ã»ÓÐ¾ÍÑ¯ÎÊÊÇ·ñ½øÈëERP²éÑ¯
    If Dir(tsvFilePath) = "" Then

        Dim resultL As VbMsgBoxResult
        resultL = MsgBox("ÔÚ¹¤×÷Ä¿Â¼ÏÂÎ´ÕÒµ½ºÏ·¨µÄÅúÁ¿²éÑ¯½á¹ûÎÄ¼þ£¡" & vbCrLf & vbCrLf & vbCrLf & _
                  "ÊÇ·ñ´ò¿ªERPÏµÍ³½øÐÐÅúÁ¿²éÑ¯£¿" & vbCrLf & vbCrLf & _
                  "ÊÇ-µÇÂ¼ERPÏµÍ³¿ªÊ¼²éÑ¯" & vbCrLf & vbCrLf & _
                  "·ñ-Ñ¡ÔñTSVÎÄ¼þÂ·¾¶" & vbCrLf, _
                  vbQuestion + vbMsgBoxSetForeground + vbYesNoCancel)

        If resultL = vbYes Then

            AutoLoginERP "RD_ENGINEER", "123456"
            'FindERP
            Exit Sub

        ElseIf resultL = vbNo Then

            CommonDialog1.FileName = ""
            CommonDialog1.DialogTitle = "ÇëÑ¡Ôñ.tsvÎÄ¼þ"
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
                MsgBox "±ØÐëÎª.tsvÎÄ¼þ£¡", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "¾¯¸æ"
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

    'µ¼ÈëtsvÎÄ¼þÄÚÐÅÏ¢
    ImportTSV

    '×ª»»BMFÎÄ¼þ¸ñÊ½±ÜÃâ³öÏÖÂÒÂë
    Process 75, "×Ô¶¯×ª»»BMFÎÄ¼þ¸ñÊ½ÎªANSI ..."
    BmfToAnsi

    '²éÕÒbmfÎÄ¼þÖÐÓÐÁÏºÅ£¬µ«ÊÇÃ»ÓÐÎïÁÏÃèÊöµÄÐÐ
    '¸ø³öÌáÊ¾ÊÇ·ñ×Ô¶¯ÁªÍø¸üÐÂÎïÁÏÃèÊö
    '±¾¹¦ÄÜÈ·¶¨¿ÉÒÔÊµÏÖ µ«Î´Íê³É
    'GetInfoFromERP
    'GetInfoFromERP "RD_ENGINEER", "123456"

    'Ö±½Ó½øÈëµÚ2½×¶Î Éú³ÉExcel¸ñÊ½BOM½×¶Î
    BomStage_Two

End Sub


'*************************************************************************
'**º¯ Êý Ãû£ºBomStage_Two
'**Êä    Èë£ºÎÞ
'**Êä    ³ö£ºÎÞ
'**¹¦ÄÜÃèÊö£º¸ù¾ÝCheckBoxµÄ×´Ì¬´´½¨ExcelÎÄ¼þºÍÉú³ÉÏàÓ¦µÄBOM
'            Á÷³ÌÈçÏÂ:
'            1.¸ù¾ÝÄ£°æ´´½¨Excel BOM
'            2.¸ù¾ÝÐèÒªµ÷ÕûExcel ¸ñÊ½
'            3.¶ÁÈ¡bmf(BOM Maker File)ÎÄ¼þ ½«ÐÅÏ¢ÌîÈëExcel
'            4.¸ù¾ÝÐÅÏ¢µ÷ÕûExcel¸ñÊ½
'            5.É¨ÃèExcel¸ñÊ½ ÐÞÕý²¿·Ö¸ñÊ½
'            6.Íê³É
'            ×¢Òâ£ºÁìÁÏBOMÖÐµÄ¿â´æÐÅÏ¢±ØÐë±£Ö¤ÊÇ×îÐÂµÄ¡£
'                  Òò´Ë³ÌÐò»á¼ì²étsvÎÄ¼þµÄ²úÉúÊ±¼ä
'                  Ê±¼ä²»ÔÚÈýÌìÄÚ»áÌáÊ¾£¬ÖØÐÂ²éÑ¯ÿ
'**È«¾Ö±äÁ¿£º
'**µ÷ÓÃÄ£¿é£º
'**×÷    Õß£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-22 12:11:09
'**ÐÞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**°æ    ±¾£ºV3.2.38
'*************************************************************************
Public Sub BomStage_Two()
    
    'ÓÉÄ£°å´´½¨ExcelÎÄ¼þ ²¢Éú³ÉÐèÒªµÄExcel BOMÎÄ¼þ
    BomCreate
    
     '============================================
    'ÏÔÊ¾½á¹ûÐÅÏ¢ Ôª¼þÊýÁ¿¸öÊý
    '============================================
    'PartNum(0) : NcPartNum
    'PartNum(1) : DbgPartNum
    'PartNum(2) : DbNcPartNum
    
    'PartNum(3) : LeadPartNum
    'PartNum(4) : SmtPartNum
    'PartNum(5) : OtherPartNum
    '============================================
    
    Dim msgstr As String
    msgstr = "             BOM ÎÄ¼þ´´½¨³É¹¦£¡" & vbCrLf & vbCrLf
    msgstr = msgstr + "          ²å×°   Ôª¼þ¸öÊýÎª   £º " & PartNum(3) & vbCrLf
    msgstr = msgstr + "          Ìù×°   Ôª¼þ¸öÊýÎª   £º " & PartNum(4) & vbCrLf
    msgstr = msgstr + "          ÆäËû   Ôª¼þ¸öÊýÎª   £º " & PartNum(5) & vbCrLf & vbCrLf
    
    msgstr = msgstr + "          NC     Ôª¼þ¸öÊýÎª   £º " & PartNum(0) & vbCrLf
    msgstr = msgstr + "          DBG    Ôª¼þ¸öÊýÎª   £º " & PartNum(1) & vbCrLf
    msgstr = msgstr + "          DBG_NC Ôª¼þ¸öÊýÎª   £º " & PartNum(2) & vbCrLf & vbCrLf
    msgstr = msgstr + "          Éú³ÉµÄbmfÎÄ¼þ²»½¨ÒéÊÖ¶¯ÐÞ¸Ä" & vbCrLf & vbCrLf
    msgstr = msgstr + "    ×¢Òâ£ºÉú³ÉµÄBOMÎÄ¼þÐèÒª¼ì²éÐÞ¸Äºó²Å¿É¹©ÆÀÉó "
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "BOM ÎÄ¼þ´´½¨³É¹¦"
    
    '´ò¿ªÉú³ÉµÄBOMÄ¿Â¼
    ShellExecute 0, "open", ProjectDir & "\BOM", "", "", 1
    
    'Çå¿Õ³ÌÐòÒÀÀµµÄÂ·¾¶ÐÅÏ¢£¬±ãÓÚÏÂ´Î×ª»»¿ªÊ¼
    ClearPath
    
    Process 100, "Íê³É£¡"
    
    Exit Sub

ErrorHandle:
    
    'É¾³ýÎ´ÄÜ³É¹¦Éú³ÉµÄÎÄ¼þ
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
        'tsvÎÄ¼þÊÇ·ñÊ§Ð§£¿
        If DateDiff("d", CDate(GetFileWriteTime(tsvFilePath)), Now) > 3 Then
            MsgBox "tsvÎÄ¼þÒÑ¾­¹ýÆÚ[" & GetFileWriteTime(tsvFilePath) & "]£¡" & vbCrLf & vbCrLf & _
                   "Éú³ÉÁìÁÏBOMÐè×îÐÂµÄtsvÎÄ¼þ£¡", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "¾¯¸æ"
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
