VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateLib 
   Caption         =   "���ļ�����"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   4920
   LinkTopic       =   "MakeLib"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6060
   ScaleWidth      =   4920
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5953
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
      Caption         =   "��˸���"
      Height          =   615
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   180
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "����˵����"
      Height          =   555
      Left            =   300
      TabIndex        =   1
      Top             =   1020
      Width           =   4215
   End
   Begin VB.Menu menuMountTpye 
      Caption         =   "��װ����"
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
Attribute VB_Name = "frmUpdateLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oldlibfile As String
Public newlibfile As String

Private Sub UpdateSTD()
    '��ȡ�ɿ��ļ� �¿��ļ� ���ɿ��ļ���ͬ��װ�ĵ���װ��Ϣת�Ƶ��¿��ļ���
    '�����¿��ļ���ʾ�����µķ�װ��Ӧ����װ����
    
    If oldlibfile = "" Or newlibfile = "" Then
        Exit Sub
    End If
    
    Dim oldLib                As String
    Dim newLib                As String
    Dim oldLibLine()          As String
    Dim newLibLine()          As String
    Dim tmpLibLine()          As String
    
    Dim oldAtom()             As String
    Dim newAtom()             As String
    
    newLib = GetFileContents(newlibfile)
    
    oldLibLine = OpenLibs(oldlibfile)
    
    newLibLine = Split(newLib, vbCrLf)
    tmpLibLine = Split(newLib, vbCrLf)
    
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
    Dim k As Integer
    
    For k = 2 To UBound(tmpLibLine) - 1
        Do While InStr(tmpLibLine(k), Space(2))
            tmpLibLine(k) = Replace(tmpLibLine(k), Space(2), Space(1)) '�������Ŀո�
        Loop
    Next k
    
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
    
    '����Ƿ����еķ�װ���ж�Ӧ����װ���ͣ�û�еĻ���ʾ���
    For j = 2 To UBound(tmpLibLine) - 1
        newAtom = Split(tmpLibLine(j), Space(1))
        If UBound(newAtom) = 3 Then
            
        ElseIf UBound(newAtom) = 2 Then
            '�����װ����
            Dim inputstr As String
            Dim typeStr  As String
            
INPUTVALUE:
            inputstr = "�������װ��" & newAtom(0) & "����Ӧ����װ����" & vbCrLf & vbCrLf
            inputstr = inputstr & "L �� ��װԪ��" & vbCrLf
            inputstr = inputstr & "S �� ��װԪ��" & vbCrLf
            inputstr = inputstr & "S+�� ��װԪ��&��װԪ��" & vbCrLf
            inputstr = inputstr & "N �� NoneԪ��" & vbCrLf

            typeStr = InputBox(inputstr, "ѡ����װ����")

            If typeStr = "L" Or typeStr = "S" Or typeStr = "S+" Or typeStr = "N" Then
                tmpLibLine(j) = tmpLibLine(j) & Space(1) & typeStr
                newLibLine(j) = newLibLine(j) & Space(10) & typeStr
            ElseIf typeStr = "" Then
                MsgBox "���ֶ��򿪿��ļ���дδ��ɵķ�װ����!", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
                Exit For
            Else
                MsgBox "����������������룡", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
                GoTo INPUTVALUE
            End If
        Else
            MsgBox "����Ŀ��ļ��Ǵ���ģ��ܾ����¿��ļ���", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "����"
            Exit Sub
        End If
    Next j
    
    If Dir(oldlibfile) <> "" Then
        Kill oldlibfile
    End If
    
    Open oldlibfile For Binary Access Write As #1
    Seek #1, 1
    Put #1, , newLibLine(0) & vbCrLf
    Put #1, , newLibLine(1) & vbCrLf
    Put #1, , vbCrLf
    
    For j = 2 To UBound(newLibLine) - 1
        Put #1, , newLibLine(j) & vbCrLf
    Next j
        
    Put #1, , vbCrLf

    Close #1
        
    'MsgBox "Done!", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "��ʾ"
    '�򿪿��ļ�
    Shell "notepad " & LibFilePath, vbMaximizedFocus
    
    Unload Me
ErrorHandle:

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
        newlibfile = filePath
        oldlibfile = App.Path & "\STD.lst"
        UpdateSTD
    End If
End Sub

Private Sub Command1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
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
        newlibfile = filePath
        oldlibfile = App.Path & "\STD.lst"
        UpdateSTD
    End If
End Sub

Private Sub Form_Load()
    SetWindowsPos_TopMost Me.hwnd
    
    Dim msgstr As String
    msgstr = Label1.Caption & vbCrLf & vbCrLf
    msgstr = msgstr + "����1��ѡ���Power PCB�е�����STD.lst��" & vbCrLf & vbCrLf
    msgstr = msgstr + "����2��ѡ��λ��SMB Group�汾���п��ļ���" & vbCrLf & vbCrLf
    
    msgstr = msgstr + "ע�⣺" & vbCrLf & vbCrLf
    msgstr = msgstr + "ʹ�÷��� 1 ���뽫���ɵĿ��ļ��ύ���汾���С�" & vbCrLf
    
    Label1.Caption = msgstr
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Command1_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
    With ListView1
        If (ColumnHeader.Index - 1) = .SortKey Then
            .SortOrder = (.SortOrder + 1) Mod 2
            .Sorted = True
        Else
            .Sorted = False
            .SortOrder = 0
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
        End If
    End With
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Dim j As Long, i As Long
    If Button = 1 Then
        If ListView1.HitTest(x, y) Is Nothing Then
            Exit Sub
        End If
        
        j = ListView1.HitTest(x, y).Index
        ListView1.ListItems(j).Selected = True
        
        For i = 0 To 3
            menuTypeSub(i).Checked = False
        Next
        
        Select Case List1.SelectedItem.SubItems(2)
            Case "L": menuTypeSub(0).Checked = True
            Case "S": menuTypeSub(1).Checked = True
            Case "S+": menuTypeSub(2).Checked = True
            Case "N": menuTypeSub(3).Checked = True
        End Select
        
        If menuMountTpye.Enabled = True Then
            PopupMenu menuMountTpye
        End If
        
    End If
End Sub

Private Sub menuTypeSub_Click(Index As Integer)
On Error Resume Next
    Dim PID As Long, rtn As Long
    PID = CLng(List1.SelectedItem.SubItems(1))
    If mnuSetProClassSub(Index).Checked = True Then Exit Sub
    Select Case Index
    Case 0: rtn = SetProClass(PID, REALTIME_PRIORITY_CLASS)
    Case 1: rtn = SetProClass(PID, HIGH_PRIORITY_CLASS)
    Case 2: rtn = SetProClass(PID, 32768)
    Case 3: rtn = SetProClass(PID, NORMAL_PRIORITY_CLASS)
    Case 4: rtn = SetProClass(PID, 16384)
    Case 5: rtn = SetProClass(PID, IDLE_PRIORITY_CLASS)
    End Select
    If rtn = 0 Then MsgBox "�޷�Ϊ���� " & List1.SelectedItem.Text & " �������ȼ���", vbCritical
End Sub
