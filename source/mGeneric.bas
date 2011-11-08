Attribute VB_Name = "mGeneric"
'*************************************************************************************
'**模 块 名：mGeneric
'**说    明：TP-LINK SMB Switch Product Line Hardware Group 版权所有2011 - 2012(C)
'**创 建 人：Shenhao
'**日    期：2011-10-31 23:36:58
'**修 改 人：
'**日    期：
'**描    述：通用模块库
'**版    本：V3.6.3
'*************************************************************************************
Option Explicit


Public ProcInfo As StatusBar

'*************************************************************************
'**函 数 名：Process
'**输    入：ProcessNum(Integer) -
'**        ：ProcessMsg(String)  -
'**输    出：无
'**功能描述：程序执行进度 允许其他模块更新执行进度
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:00:01
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function Process(ProcessNum As Integer, ProcessMsg As String)

    MainForm.StatusBar1.Panels(1) = ProcessMsg
    MainForm.StatusBar1.Panels(2) = ProcessNum & "%"
End Function

'*************************************************************************
'**函 数 名：SetWindowsPos_TopMost
'**输    入：hwnd(Long) -
'**输    出：无
'**功能描述：将 窗口设定成永远保持在最上层
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:00:17
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function SetWindowsPos_TopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Function

'*************************************************************************
'**函 数 名：SetWindowsPos_NoTopMost
'**输    入：hwnd(Long) -
'**输    出：无
'**功能描述：取消最上层设定
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:00:33
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function SetWindowsPos_NoTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Function


'*************************************************************************
'**函 数 名：SetRegValue
'**输    入：AppName(String)            -
'**        ：KeyName(String)            -
'**        ：ByVal lType(enumRegSzType) -
'**        ：                           -
'**输    出：(Boolean) -
'**功能描述：设置注册表子项
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:00:50
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Public Function SetRegValue(AppName As String, KeyName As String, ByVal lType As enumRegSzType, ByVal KeyValue) As Boolean

     SetRegValue = SetValue(iHKEY_CURRENT_USER, "SOFTWARE\" + AppName, KeyName, lType, KeyValue)
     
End Function


'*************************************************************************
'**函 数 名：SetRegValue
'**输    入：AppName(String)            -
'**        ：KeyName(String)            -
'**        ：ByVal lType(enumRegSzType) -
'**        ：                           -
'**输    出：(Boolean) -
'**功能描述：读取注册表子项
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:00:50
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Public Function GetRegValue(AppName As String, KeyName As String, Optional DefaultKeyValue As Variant) As Variant

    If GetValue(iHKEY_CURRENT_USER, "SOFTWARE\" + AppName, KeyName, GetRegValue) = False Then
        GetRegValue = DefaultKeyValue
    End If
    
End Function

Public Function GetExcelVer() As String

    Dim sKeyReg() As String
    Dim tmpReg()  As String
    Dim Cnt       As Long
    
    
    If RegEnumKeyVal(iHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office", sKeyReg) = False Then
        GetExcelVer = ""
        Exit Function
    End If

    For Cnt = 0 To UBound(sKeyReg)
        Select Case sKeyReg(Cnt)
        Case "14.0", "12.0", "11.0", "10.0", "9.0", "8.0", "5.0", "4.0", "3.0" 'Excel各个版本号
            If RegEnumKeyVal(iHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\" & sKeyReg(Cnt) & "\Excel", tmpReg) = True Then
                GetExcelVer = sKeyReg(Cnt)
                Exit For
            End If

        Case Else
            If CInt(sKeyReg(Cnt)) > 14 Then
                GetExcelVer = sKeyReg(Cnt)
            Else
                GetExcelVer = ""
            End If
        End Select
    Next
    
End Function

'去除缓冲区中多余的Chr(0)
Public Function StripTerminator(sInput As String) As String
  Dim ZeroPos As Integer
  '找到第一个Chr(0)
  ZeroPos = InStr(1, sInput, vbNullChar)
  If ZeroPos > 0 Then '如果存在,则去掉后面所有的内容
    StripTerminator = Left$(sInput, ZeroPos - 1)
  Else '如果不存在,不做任何操作
    StripTerminator = sInput
  End If
End Function



'*************************************************************************
'**函 数 名：KillExcel
'**输    入：ExcelFilePath(String) -
'**输    出：无
'**功能描述：删除旧的Excel格式文件
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:01:29
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function KillExcel(ExcelFilePath As String)
On Error GoTo ErrorHandle
    If Dir(ExcelFilePath) <> "" Then
        Kill ExcelFilePath
    End If
    
    Exit Function
    
ErrorHandle:
    MsgBox "文件：" & vbCrLf & Right(ExcelFilePath, Len(ExcelFilePath) - InStrRev(ExcelFilePath, "\")) & vbCrLf & vbCrLf & _
           "已经打开或被占用，请将其关闭后重新运行程序！", vbCritical + vbOKOnly + vbMsgBoxSetForeground, "错误"
    End
End Function


'*************************************************************************
'**函 数 名：GetPath
'**输    入：Promt(String) - 提示内容
'**输    出：结果为选择的目录
'**功能描述：选择目录路径
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:01:41
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
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

'*************************************************************************
'**函 数 名：GetFileContents
'**输    入：filePath(String) -
'**输    出：(String) -
'**功能描述：获取文件内容 并且去掉无意义空行
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:02:21
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
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

    cleanContents = UEFLoadTextFile(filePath, UEF_ANSI)
    
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


'*************************************************************************
'**函 数 名：GetBomContents
'**输    入：filePath(String) -
'**输    出：(String) -
'**功能描述：获取物料BOM文件内容
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:02:30
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function GetBomContents(filePath As String) As String
    
    Dim BomLine()    As String
    Dim BomAtom()    As String
    Dim BomInfo      As String
    Dim AtomNum      As Integer
    Dim j            As Integer
    
    BomInfo = GetFileContents(filePath)
    
    '对于BOM文件来说需要整理下格式
    Do While InStr(1, BomInfo, vbCrLf + vbTab)
        BomInfo = Replace(BomInfo, vbCrLf + vbTab, vbTab)
    Loop
    
    '获取BOM文件的列数
    BomLine = Split(BomInfo, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组
    AtomNum = UBound(Split(BomLine(0), vbTab)) '均以首行列数为标准
    
    BomInfo = ""
    '核对每列的元素个数 元素个数不足 尝试补齐
     For j = 0 To UBound(BomLine) - 1
        BomAtom = Split(BomLine(j), vbTab)
        
        '元素个数不足 与下一列合并
        If UBound(BomAtom) <> AtomNum Then
            BomInfo = BomInfo + BomLine(j) + BomLine(j + 1) + vbCrLf
            j = j + 1
        Else
            BomInfo = BomInfo + BomLine(j) + vbCrLf
        End If
    Next j
    
    BomInfo = BomInfo + BomLine(j)
    
    '去掉最后的换行
    'BomInfo = Left(BomInfo, Len(BomInfo) - 2)
    
    GetBomContents = BomInfo
    
'for debug
'    If Dir(SaveAsPath & "_test.txt") <> "" Then
'        Kill SaveAsPath & "_test.txt"
'    End If
'
'    Open SaveAsPath & "_test.txt" For Output As #1
'    Print #1, GetBomContents
'    Close #1
    
End Function


'*************************************************************************
'**函 数 名：WriteTxt
'**输    入：strSourceFile(String) - 要写入的文件地址
'**        ：intRow(Long)          - 修改的行数
'**        ：StrLineNew(String)    - 写入或替换的字符串
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:02:52
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Public Function WriteTxt(strSourceFile As String, intRow As Long, StrLineNew As String)

    Dim StrOut As String, tmpStrLine As String
    Dim X As Long
    If Dir(strSourceFile) <> "" Then
        Open strSourceFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpStrLine
            X = X + 1
            If X = intRow Then tmpStrLine = StrLineNew
            If Not EOF(1) Then
                StrOut = StrOut & tmpStrLine & vbCrLf
            Else
                StrOut = StrOut & tmpStrLine
            End If
            'Debug.Print x
        Loop
        Close #1
    Else
        StrOut = StrLineNew
    End If
    
    Open strSourceFile For Output As #1
    Print #1, StrOut
    Close #1

End Function

'*************************************************************************
'**函 数 名：ReadTxt
'**输    入：StrFile(String) -  文件地址
'**        ：intRow(Long)    -  读取的行数
'**输    出：(String) -  要输出的文本
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:03:13
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Public Function ReadTxt(StrFile As String, intRow As Long) As String
    Dim StrOut As String, tmpStrLine As String
    Dim X As Long
    
    If Dir(StrFile, vbNormal) <> "" Then
        Open StrFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, tmpStrLine
            X = X + 1
            If X = intRow Then ReadTxt = tmpStrLine: Exit Do
        Loop
        Close #1
    End If
    
End Function

'*************************************************************************
'**函 数 名：AutoLoginERP
'**输    入：uid(String) -
'**        ：pwd(String) -
'**输    出：无
'**功能描述：自动登录ERP系统 为批量查询做准备
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-01 00:44:04
'**修 改 人：
'**日    期：
'**版    本：V3.6.7
'*************************************************************************
Function AutoLoginERP(uid As String, pwd As String)
    Dim i As Integer
    Dim j As Integer
    Dim orgName As String
    
    orgName = "内销 研发－工程师"
    
    '获取库存类型
    Dim SelStorage As String
    
    SelStorage = GetRegValue(App.EXEName, "Storage", "TP1")
    
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


'*************************************************************************
'**函 数 名：GetInfoFromERP
'**输    入：uid(String) -
'**        ：pwd(String) -
'**输    出：无
'**功能描述：在ERP系统中获取信息，本Function未完成
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-01 00:44:33
'**修 改 人：
'**日    期：
'**版    本：V3.6.7
'*************************************************************************
Function GetInfoFromERP(uid As String, pwd As String)
    Dim i As Integer
    Dim j As Integer
    Dim orgName As String
    
    orgName = "内销 研发－工程师"
    
    '获取库存类型
    Dim SelStorage As String
    
    SelStorage = GetRegValue(App.EXEName, "Storage", "TP1")
    
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
        .Visible = True
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
                            If InStr(.Document.All(j).innerText, "替代料查询") <> 0 Then
                                .Document.All(j).Click
                                GoTo Check_Done
                            End If
                        End If
                    Next j
                    
                End If
            End If
        Next i
        
Check_Done:
        '进入到替代料查询位置 上传Excel 批量查询替代料
        Do While .busy Or .ReadyState <> 4
        Loop
        
        '<option value="222">TP4</option>
        '<option value="262">TP5</option>
        '<option value="382">TP7</option>
        '<option value="121">TP1</option>
        '<option value="122">TP2</option>
        '<option value="123">TP3</option>
        '<option value="442">TP8</option>
        '<option value="321">TP6</option>
        .Document.Forms(0).All("OrganizationCode").Value = "121"
        '<option value="L">批量物料</option>
        '<option value="S">单个物料</option>
        .Document.GetElementById("QueryType").Value = "L"
        .Document.GetElementById("QueryType").FireEvent "onchange"
        .Document.GetElementById("QueryType").FireEvent "onclick"
        .Document.GetElementById("QueryType").FireEvent "onblur"
        
        '进入到替代料查询位置 上传Excel 批量查询替代料
        Do While .busy Or .ReadyState <> 4
        Loop
        
        'VB.Clipboard.SetText "E:\其他任务\MakeBomTest\test\TL-SL1210-2.0\BOM\TL-SL1210(UN)_2.0_SCHV_2.0_DEV1_PCBV_1.0SP1_批量资源查询.xls"
        
        .Document.GetElementById("UploadFile_oafileUpload").Focus
        .Document.GetElementById("UploadFile_oafileUpload").Click
        DoEvents
        
        '.Document.Forms(0).All("UploadFile_oafileUpload").Value = "E:\其他任务\MakeBomTest\test\TL-SL1210-2.0\BOM\TL-SL1210(UN)_2.0_SCHV_2.0_DEV1_PCBV_1.0SP1_批量资源查询.xls"
        
        .Document.GetElementById("Go").FireEvent "onclick"
        
        '进入到替代料查询位置 上传Excel 批量查询替代料
        Do While .busy Or .ReadyState <> 4
        Loop
        
        '重新开始查询
        
    End With
    
End Function


'*************************************************************************
'**函 数 名：GetFileCreateTime
'**输    入：filePath(String) -
'**输    出：(String) -
'**功能描述：获取文件建立时间
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:03:38
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function GetFileCreateTime(filePath As String) As String

    Dim FileHandle As Long
    Dim fileinfo As BY_HANDLE_FILE_INFORMATION
    Dim lpReOpenBuff As OFSTRUCT, ft As SYSTEMTIME
    Dim tZone As TIME_ZONE_INFORMATION
    
    Dim dtCreate As Date ' 建立时间。

    Dim bias As Long
    ' 先取得文件的Handle。
    FileHandle = OpenFile(filePath, lpReOpenBuff, OF_READ)
    ' 利用文件的Handle读取文件信息。
    Call GetFileInformationByHandle(FileHandle, fileinfo)
    Call CloseHandle(FileHandle)
    ' 读取时间信息， 因为上一步骤的文件时间是格林威治时间。
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias ' 时间差， 以分为单位。
    Call FileTimeToSystemTime(fileinfo.ftCreationTime, ft) ' 转换时间结构。
    dtCreate = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)
    
    GetFileCreateTime = CStr(dtCreate)

End Function


'*************************************************************************
'**函 数 名：GetFileWriteTime
'**输    入：filePath(String) -
'**输    出：(String) -
'**功能描述：获取文件修改时间
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:03:46
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function GetFileWriteTime(filePath As String) As String

    Dim FileHandle As Long
    Dim fileinfo As BY_HANDLE_FILE_INFORMATION
    Dim lpReOpenBuff As OFSTRUCT, ft As SYSTEMTIME
    Dim tZone As TIME_ZONE_INFORMATION
    
    Dim dtWrite As Date ' 修改时间。
    Dim bias As Long
    ' 先取得文件的Handle。
    FileHandle = OpenFile(filePath, lpReOpenBuff, OF_READ)
    ' 利用文件的Handle读取文件信息。
    Call GetFileInformationByHandle(FileHandle, fileinfo)
    Call CloseHandle(FileHandle)
    ' 读取时间信息， 因为上一步骤的文件时间是格林威治时间。
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias ' 时间差， 以分为单位。
    
    Call FileTimeToSystemTime(fileinfo.ftLastWriteTime, ft)
    dtWrite = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)

    GetFileWriteTime = CStr(dtWrite)

End Function


'*************************************************************************
'**函 数 名：GetFileAccessTime
'**输    入：filePath(String) -
'**输    出：(String) -
'**功能描述：获取文件存取时间
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-11-07 22:03:55
'**修 改 人：
'**日    期：
'**版    本：V3.6.16
'*************************************************************************
Function GetFileAccessTime(filePath As String) As String

    Dim FileHandle As Long
    Dim fileinfo As BY_HANDLE_FILE_INFORMATION
    Dim lpReOpenBuff As OFSTRUCT, ft As SYSTEMTIME
    Dim tZone As TIME_ZONE_INFORMATION
    
    Dim dtAccess As Date ' 存取日期。
    Dim bias As Long
    ' 先取得文件的Handle。
    FileHandle = OpenFile(filePath, lpReOpenBuff, OF_READ)
    ' 利用文件的Handle读取文件信息。
    Call GetFileInformationByHandle(FileHandle, fileinfo)
    Call CloseHandle(FileHandle)
    ' 读取时间信息， 因为上一步骤的文件时间是格林威治时间。
    Call GetTimeZoneInformation(tZone)
    bias = tZone.bias ' 时间差， 以分为单位。
    
    Call FileTimeToSystemTime(fileinfo.ftLastAccessTime, ft)
    dtAccess = DateSerial(ft.wYear, ft.wMonth, ft.wDay) + TimeSerial(ft.wHour, ft.wMinute - bias, ft.wSecond)

    GetFileAccessTime = CStr(dtAccess)

End Function

