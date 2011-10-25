VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLib 
   Caption         =   "��װ��"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "UpdateLib.frx":0000
   LinkTopic       =   "MakeLib"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7050
   ScaleWidth      =   9195
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   8595
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10186
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
      Caption         =   "����"
      Height          =   615
      Left            =   480
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
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
      Caption         =   "�汾��Ϣ"
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.Menu menuMountTpye 
      Caption         =   "��װ����"
      Visible         =   0   'False
      Begin VB.Menu menuTypeSub 
         Caption         =   "S����װԪ��"
         Index           =   0
      End
      Begin VB.Menu menuTypeSub 
         Caption         =   "L����װԪ��"
         Index           =   1
      End
      Begin VB.Menu menuTypeSub 
         Caption         =   "S+����װԪ��"
         Index           =   2
      End
      Begin VB.Menu menuTypeSub 
         Caption         =   "N��NoneԪ��"
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
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'�����
Public Enum MenuType
MENU_S = 0   'SMD Type
MENU_L       'LEAD Type
MENU_SP      'SMD Type
MENU_N       'None Type
End Enum

Private Sub UpdateSTD(newlibfile As String)
    '��ȡ�ɿ��ļ� �¿��ļ� ���ɿ��ļ���ͬ��װ�ĵ���װ��Ϣת�Ƶ��¿��ļ���
    '�����¿��ļ���ʾ�����µķ�װ��Ӧ����װ����
    
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
    
    '��Ӱ汾��Ϣ ��ָ������Ϣ
    newAtom = Split(tmpLibLine(0), Space(1))
    If UBound(newAtom) = 13 Then
        newLibLine(0) = newLibLine(0) & Space(14) & "VERSION:" & Space(2) & CStr(Now)
    End If
    
    newAtom = Split(tmpLibLine(1), Space(1))
    If UBound(newAtom) = 34 Then
        newLibLine(1) = newLibLine(1) & Space(14) & "Mount Type"
    End If
    'newLibLine(2) = newLibLine(2) ��ȡ�ļ����ݵ�ʱ��ȥ���˿���
    
    '���ɿ����Ϣ���뵽�¿� �����ɿ�
    Dim i As Integer
    Dim j As Integer
    
    For i = 2 To UBound(tmpLibLine) - 1
        Do While InStr(tmpLibLine(i), Space(2))
            tmpLibLine(i) = Replace(tmpLibLine(i), Space(2), Space(1)) '�������Ŀո�
        Loop
    Next i
    
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
    
    '������º�Ŀ��ļ� �ȴ��޸Ļ���
    LoadLibs
    
    '����װ��������
    With ListView1
        .SortOrder = 0    '˳������
        .SortKey = 3 - 1  '����װ��������
        .Sorted = True
    End With

End Sub


Private Sub Command1_Click()
    Dim filePath As String
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "��ѡ��STD.LST�ļ�"
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

Private Sub Command1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�����ϷŲ���
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

    msgstr = msgstr + "��1����ѡ���Power PCB�е�����STD.lst��" & vbCrLf & vbCrLf
    msgstr = msgstr + "��2��������Ҫ�޸ĵķ�װ��ʹ�������ѡ����ȷ�ķ�װ���͡�" & vbCrLf & vbCrLf
    
    msgstr = msgstr + "ע�⣺" & vbCrLf & vbCrLf
    msgstr = msgstr + "�����޸ĺ�Ŀ��ļ����ύ���汾���С�" & vbCrLf
    
    MsgBox msgstr, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "������Ϣ"
End Sub

Private Sub Form_Load()
    SetWindowsPos_TopMost Me.hwnd
    
    Label1.Caption = "��װ��汾��" & GetLibsVersion
    
    '��ȡ�ɿ��ļ� ��ʾ��ListView��
    InitList
    
End Sub

Private Sub InitList()

    '��ӿ��ļ����ݵ�ListView
    Dim gwidth        As Integer
    
    '��ʼ��ListView1
    gwidth = ListView1.Width / 40
    
    ListView1.ColumnHeaders.Add , , "���", gwidth * 3
    ListView1.ColumnHeaders.Add , , "��װ����", gwidth * 18
    ListView1.ColumnHeaders.Add , , "��װ����", gwidth * 6
    ListView1.ColumnHeaders.Add , , "�޸�ʱ��", gwidth * 10
    
    LoadLibs
    
End Sub

Private Sub LoadLibs()
    Dim LibLine()     As String
    Dim LibAtom()     As String
    Dim i             As Integer
    
    '����б�
    ListView1.ListItems.Clear
    
    '���Ԫ��
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
    
    UpdateListviewColor
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '����ļ���Ϣ�Ƿ���������������Ҫ��ʾ�ֶ��޸�
    
    '������ļ��������Ժ���ȷ��
    Dim LibLine()          As String
    Dim LibAtom()          As String
    Dim i                  As Integer
    
    LibLine = OpenLibs()
    For i = 2 To UBound(LibLine) - 1
        LibAtom = Split(LibLine(i), Space(1))
        If UBound(LibAtom) <> 3 Then
            If MsgBox("���ļ�STD.lst�޸Ĳ�������ȷ���˳���" & vbCrLf & vbCrLf & _
                      "��ʾ��" & vbCrLf & "�˳������ֶ��޸�Ϊ��ȷ�ļ�����������޷���ȷ���У�", _
                      vbMsgBoxSetForeground + vbQuestion + vbYesNo, "����") = vbYes Then
                      
                End
            Else
                Cancel = True
            End If
        End If
    Next i
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim j As Long, i As Long
    
    '��Ӧ������ �� �Ҽ��¼� �� �м��¼�
    If Button = vbLeftButton Or Button = vbRightButton Or Button = vbMiddleButton Then
        If ListView1.HitTest(X, Y) Is Nothing Then
            Exit Sub
        End If
        
        j = ListView1.HitTest(X, Y).Index
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
        '����Ӧ������
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

    '����Ӧ���޸�ͬ�������ļ���
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

'ListView ������ʾ��ͬ����ɫ
Private Sub UpdateListviewColor()
    '�˴����ø�����ɫ��ͬ
    'ͼ��ؼ���Ҫ�������
    Picture1.BorderStyle = vbBSNone
    Picture1.AutoRedraw = True
    'Picture1.Visible = False
    
    '�߶�Ϊ�����б�
    Picture1.Width = ListView1.Width
    Picture1.Height = ListView1.ListItems(1).Height * 2
    
    '�������м����ɫ
    Picture1.ScaleMode = vbUser
    Picture1.ScaleHeight = 2
    Picture1.ScaleWidth = 1
    Picture1.Line (0, 0)-(1, 1), vbWhite, BF
    Picture1.Line (0, 1)-(1, 2), RGB(220, 226, 241), BF
    
    '��ؼ��ĵط�
   ListView1.PictureAlignment = lvwTile
   ListView1.Picture = Picture1.Image
    
End Sub


'ListView��Ĭ�ϵ������ܶ��ǰ����ַ���˳���ŵģ�����������˳��ʱ������������У�9������10�ĺ��档
'���³���������һ����
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    
    '��ʱ�� ���Բ���ʱʹ��
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
        
        If ColumnHeader.Text = "���" Then
            
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
    
    'MsgBox "����: " & ListView1.ListItems.Count & " ��ʱ: " & GetTickCount - lngStart & "ms"
    
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
