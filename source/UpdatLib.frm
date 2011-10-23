VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUpdateLib 
   Caption         =   "库文件更新"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "MakeLib"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3015
   ScaleWidth      =   4875
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "点此更新"
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
      Caption         =   "更新说明："
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
    '读取旧库文件 新库文件 将旧库文件相同封装的的贴装信息转移到新库文件中
    '最后打开新库文件提示填入新的封装对应的贴装类型
    
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
    
    '添加版本信息 等指定的信息
    newAtom = Split(tmpLibLine(0), Space(1))
    If UBound(newAtom) = 13 Then
        newLibLine(0) = newLibLine(0) & Space(14) & "VERSION:" & Space(2) & CStr(Now)
    End If
    
    newAtom = Split(tmpLibLine(1), Space(1))
    If UBound(newAtom) = 34 Then
        newLibLine(1) = newLibLine(1) & Space(14) & "Mount Type"
    End If
    'newLibLine(2) = newLibLine(2) 获取文件内容的时候去掉了空行
    
    '将旧库的信息导入到新库 遍历旧库
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    For k = 2 To UBound(tmpLibLine) - 1
        Do While InStr(tmpLibLine(k), Space(2))
            tmpLibLine(k) = Replace(tmpLibLine(k), Space(2), Space(1)) '清除多余的空格
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
    
    '检查是否所有的封装都有对应的贴装类型，没有的话提示添加
    For j = 2 To UBound(tmpLibLine) - 1
        newAtom = Split(tmpLibLine(j), Space(1))
        If UBound(newAtom) = 3 Then
            
        ElseIf UBound(newAtom) = 2 Then
            '添加贴装类型
            Dim inputstr As String
            Dim typeStr  As String
            
INPUTVALUE:
            inputstr = "请输入封装【" & newAtom(0) & "】对应的贴装类型" & vbCrLf & vbCrLf
            inputstr = inputstr & "L ： 插装元件" & vbCrLf
            inputstr = inputstr & "S ： 贴装元件" & vbCrLf
            inputstr = inputstr & "S+： 贴装元件&插装元件" & vbCrLf
            inputstr = inputstr & "N ： None元件" & vbCrLf

            typeStr = InputBox(inputstr, "选择贴装类型")

            If typeStr = "L" Or typeStr = "S" Or typeStr = "S+" Or typeStr = "N" Then
                tmpLibLine(j) = tmpLibLine(j) & Space(1) & typeStr
                newLibLine(j) = newLibLine(j) & Space(10) & typeStr
            ElseIf typeStr = "" Then
                MsgBox "请手动打开库文件填写未完成的封装分类!", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
                Exit For
            Else
                MsgBox "输入错误！请重新输入！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
                GoTo INPUTVALUE
            End If
        Else
            MsgBox "导入的库文件是错误的，拒绝更新库文件！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
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
        
    'MsgBox "Done!", vbInformation + vbMsgBoxSetForeground + vbOKOnly, "提示"
    '打开库文件
    Shell "notepad " & LibFilePath, vbMaximizedFocus
    
    Unload Me
ErrorHandle:
    

End Sub

Private Sub Command1_Click()
    Dim filePath As String
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "请选择STD.LST文件"
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
    '允许拖放操作
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
    msgstr = msgstr + "方法1：选择从Power PCB中导出的STD.lst。" & vbCrLf & vbCrLf
    msgstr = msgstr + "方法2：选择位于SMB Group版本库中库文件。" & vbCrLf & vbCrLf
    
    msgstr = msgstr + "注意：" & vbCrLf & vbCrLf
    msgstr = msgstr + "使用方法 1 必须将生成的库文件提交到版本库中。" & vbCrLf
    
    Label1.Caption = msgstr
    
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

