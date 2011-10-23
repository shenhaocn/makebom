VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUpdateLib 
   Caption         =   "���ļ�����"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "MakeLib"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3015
   ScaleWidth      =   4875
   StartUpPosition =   2  '��Ļ����
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
      Height          =   1875
      Left            =   300
      TabIndex        =   1
      Top             =   1020
      Width           =   4215
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

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

