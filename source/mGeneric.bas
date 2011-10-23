Attribute VB_Name = "mGeneric"
Option Explicit

'==============================================================================
' Constant defining ( 常数定义 )
'==============================================================================
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
'==============================================================================
' API function declare ( API函数声明 )
'==============================================================================
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = &H2 '不更动目前视窗位置
Const SWP_NOSIZE = &H1 '不更动目前视窗大小
Const HWND_TOPMOST = -1 '设定为最上层
Const HWND_NOTOPMOST = -2 '取消最上层设定
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public ProcInfo As StatusBar

'将 窗口设定成永远保持在最上层
Function SetWindowsPos_TopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Function

'取消最上层设定
Function SetWindowsPos_NoTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Function

'获取文件内容 并且去掉无意义空行
Function GetFileContents(filePath As String) As String
    Dim cfgfile            As String
    Dim cfgfilenum         As Integer
    Dim cfgfileContents    As String
    Dim cfgfileInfo()      As String
    Dim cleanContents      As String
    
    cfgfile = filePath
    
    cfgfilenum = FreeFile
    Open cfgfile For Binary As #cfgfilenum
        cfgfileContents = Space(LOF(cfgfilenum))
        Get #cfgfilenum, , cfgfileContents
    Close cfgfilenum

    cleanContents = cfgfileContents
    
    '去掉空行
    '删除在文件中读取的单独的无意义的空行
    Do While InStr(cleanContents, " " + vbCrLf) > 0
        '清除换行符前的空格
        cleanContents = Replace(cleanContents, " " + vbCrLf, vbCrLf)
    Loop
    Do While InStr(1, cleanContents, vbCrLf + vbCrLf)
        '清除行与行之间的空行
        cleanContents = Replace(cleanContents, vbCrLf + vbCrLf, vbCrLf)
    Loop
    If InStr(cleanContents, vbCrLf) = 1 Then
        '清除为首的空行
        cleanContents = Replace(cleanContents, vbCrLf, "", 1, 1)
    End If
    
    GetFileContents = cleanContents
End Function

'参数为提示内容，返回结果为选择的目录
Function GetPath(Promt As String)
    Dim objshell
    Dim objfolder
    Set objshell = CreateObject("Shell.Application")
        Set objfolder = objshell.BrowseForFolder(0, Promt, 0, 0)
            If Not objfolder Is Nothing Then
                GetPath = objfolder.self.Path
            End If
        Set objfolder = Nothing
    Set objshell = Nothing
End Function

'程序执行进度 允许其他模块更新执行进度
Function Process(ProcessNum As Integer, ProcessMsg As String)

    MainForm.StatusBar1.Panels(1) = ProcessMsg
    MainForm.StatusBar1.Panels(2) = ProcessNum & "%"
End Function

Function KillExcel(ExcelFilePath As String)
    If Dir(ExcelFilePath) <> "" Then
        Kill ExcelFilePath
    End If
End Function

Function AutoLoginERP(uid As String, pwd As String)
    Dim i As Integer
    Dim j As Integer
    Dim orgName As String
    
    orgName = "内销 研发－工程师"
    
    '获取库存类型
    Dim SelStorage As String
    
    SelStorage = GetSetting(App.EXEName, "SelectStorage", "库存类型", "TP1")
    
    Select Case SelStorage
    Case "TP1"
        orgName = "内销 研发－工程师"
    Case "TP2"
        orgName = "外销 研发－工程师"
    Case "TP3"
        orgName = "OEM 研发－工程师"
    Case Else
        
        orgName = "内销 研发－工程师"
    End Select
        
    'MsgBox "正在打开ERP系统，请稍等..."
    
    With CreateObject("InternetExplorer.Application")
        .Visible = False
        .Navigate "http://erpprod.tplink.net:8007/OA_HTML/AppsLocalLogin.jsp?cancelUrl=/OA_HTML/AppsLocalLogin.jsp&langCode=ZHS"
        Do Until .ReadyState = 4
            DoEvents
        Loop
        .Document.Forms(0).All("username").Value = uid
        .Document.Forms(0).All("password").Value = pwd
        .Document.Forms(0).submit
        '自动登录结束
        
        Do While .busy Or .ReadyState <> 4
        Loop

        For i = 0 To .Document.All.Length - 1
            If (LCase(.Document.All(i).tagname)) = "a" Then
                If InStr(.Document.All(i).innerText, orgName) <> 0 Then
                    .Document.All(i).Click
                    
                    Do While .busy Or .ReadyState <> 4
                    Loop
                    
                    '进入到批量查询
                    For j = 0 To .Document.All.Length - 1
                        If (LCase(.Document.All(j).tagname)) = "a" Then
                            If InStr(.Document.All(j).innerText, "批量可用资源查询") <> 0 Then
                                .Document.All(j).Click
                                Exit Function
                            End If
                        End If
                    Next j
                    
                End If
            End If
        Next i
        
    End With
    
End Function

Function FindERP()
    Dim ParentWnd As Long   '父窗口句柄
    Dim ClientWnd As Long   '子窗口句柄
    
    Dim msgstr    As String
    
    '以下是取得你指定窗口句柄过程，注意修改类名和窗口名
    ParentWnd = FindWindow("SunAwtFrame", "Oracle 应用产品 - PROD")
    If ParentWnd = 0 Then
        MsgBox "没有找到父窗口", 16, "错误"
        Exit Function
    End If
    
    '取得第一个子窗口的句柄
    ClientWnd = GetWindow(ParentWnd, GW_CHILD)
    If ClientWnd = 0 Then
        MsgBox "在指定窗口中没有发现子窗口的存在", 16, "错误"
        Exit Function
    End If
    
    '开始循环查找所有相同层次的子窗口
    Do
        DoEvents
        msgstr = msgstr & "子窗口：" & ClientWnd & vbCrLf
        ClientWnd = GetWindow(ClientWnd, GW_HWNDNEXT)
    Loop While ClientWnd <> 0
    
    MsgBox "完成处理。", 64, "提示"
    
End Function

