Attribute VB_Name = "mSortLib"
'*************************************************************************************
'**模 块 名：mSortLib
'**说    明：TP-LINK SMB Switch Product Line Hardware Group 版权所有2011 - 2012(C)
'**创 建 人：Shenhao
'**日    期：2011-10-31 23:38:03
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V3.6.3
'*************************************************************************************
Option Explicit

'分类库
Public Enum MountType
LIB_LEAD = 0 'LEAD Type
LIB_SMD      'smd type
LIB_NONE     'None type
End Enum

Public LibFilePath     As String  '库文件路径信息


'*************************************************************************
'**函 数 名：InitLib
'**输    入：filePath(String) -
'**输    出：(Boolean) -
'**功能描述：初始化封装库
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-31 23:41:31
'**修 改 人：
'**日    期：
'**版    本：V3.6.3
'*************************************************************************
Function InitLib(filePath As String) As Boolean
    '是否存在库文件
    If Dir(filePath) = "" Then
         MsgBox "库文件不存在！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "严重错误"
         InitLib = False
         Exit Function
    End If
    
    LibFilePath = filePath
    
    '检验库文件的完整性和正确性
    Dim LibLine()          As String
    Dim LibAtom()          As String
    Dim i                  As Integer
    
    LibLine = OpenLibs()
    For i = 2 To UBound(LibLine) - 1
        LibAtom = Split(LibLine(i), Space(1))
        If UBound(LibAtom) <> 3 Then
            MsgBox "库文件STD.lst不完整或错误！请更新STD.lst库文件！", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "错误"
            End
        End If
    Next i
    
    InitLib = True
End Function


'*************************************************************************
'**函 数 名：OpenLibs
'**输    入：无
'**输    出：字符串数组是以一个空格分割各个属性的字符串 以各行内容为数组
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-31 23:42:02
'**修 改 人：
'**日    期：
'**版    本：V3.6.3
'*************************************************************************
Function OpenLibs() As String()
    Dim FileContents    As String
    Dim fileinfo()      As String
        
    FileContents = GetFileContents(LibFilePath)
    fileinfo = Split(FileContents, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组

    Dim j As Integer
    For j = 2 To UBound(fileinfo) - 1
        Do While InStr(fileinfo(j), Space(2))
            fileinfo(j) = Replace(fileinfo(j), Space(2), Space(1)) '清除多余的空格
        Loop
    Next j
    
    '返回的字符串数组是以一个空格分割各个属性的字符串 以各行内容为数组
    OpenLibs = fileinfo
        
End Function


'*************************************************************************
'**函 数 名：GetLibsVersion
'**输    入：无
'**输    出：(String) -
'**功能描述：获取封装库版本号
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-31 23:42:48
'**修 改 人：
'**日    期：
'**版    本：V3.6.3
'*************************************************************************
Function GetLibsVersion() As String
    Dim FileContents    As String
    Dim fileinfo()      As String
        
    FileContents = GetFileContents(LibFilePath)
    fileinfo = Split(FileContents, vbCrLf) '取出源文件行数，按照回车换行来分隔成数组
    
    '版本信息在第1行
    GetLibsVersion = Right(fileinfo(0), Len(fileinfo(0)) - InStrRev(fileinfo(0), "VERSION:") + 1)
    
    GetLibsVersion = Replace(GetLibsVersion, "VERSION:" & Space(2), "")
        
End Function


'*************************************************************************
'**函 数 名：ReadLibs
'**输    入：Lib(MountType) As String() -
'**输    出：无
'**功能描述：读取指定的封装库
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-31 23:43:14
'**修 改 人：
'**日    期：
'**版    本：V3.6.3
'*************************************************************************
Function ReadLibs(Lib As MountType) As String()
    Dim LibsInfo()      As String
    Dim Libstr()        As String
    LibsInfo = OpenLibs()
    
    '统计库中各中封装类型的数量
    Dim j               As Integer
    Dim leadNum         As Integer
    Dim smdNum          As Integer
    Dim otherNum        As Integer
    
    leadNum = 0
    smdNum = 0
    otherNum = 0
    For j = 2 To UBound(LibsInfo) - 1
        Libstr = Split(LibsInfo(j), " ")
        If UBound(Libstr) = 3 Then
            Select Case Libstr(3)
            Case "L"
                leadNum = leadNum + 1
            Case "S"
                smdNum = smdNum + 1
            Case "S+"
                leadNum = leadNum + 1
                smdNum = smdNum + 1
            Case "N"
                otherNum = otherNum + 1
            Case Else
                MsgBox "封装 [" & Libstr(0) & "]" & "未知分类", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "库文件错误！"
            End Select
        End If
    Next j
    
    Dim LibNum As Integer
    Dim cfgInfo() As String
    
    Select Case Lib
    Case LIB_LEAD
        LibNum = leadNum
        ReDim cfgInfo(leadNum) As String
    Case LIB_SMD
        LibNum = smdNum
        ReDim cfgInfo(smdNum) As String
    Case LIB_NONE
        LibNum = otherNum
        ReDim cfgInfo(otherNum) As String
    Case Else
        ' SHOULD NEVER BE HERE
    End Select
    
    Dim i As Integer
    
    i = 0
    For j = 2 To UBound(LibsInfo) - 1
        Libstr = Split(LibsInfo(j), " ")
        If UBound(Libstr) = 3 Then
            Select Case Libstr(3)
            Case "L"
                If Lib = LIB_LEAD Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case "S"
                If Lib = LIB_SMD Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case "S+"
                If Lib = LIB_LEAD Or Lib = LIB_SMD Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case "N"
                If Lib = LIB_NONE Then
                    cfgInfo(i) = Libstr(0)
                    i = i + 1
                End If
            Case Else
                MsgBox "封装 [" & Libstr(0) & "]" & "未知分类", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "库文件错误！"
            End Select
        End If
    Next j

    ReadLibs = cfgInfo
        
End Function


'*************************************************************************
'**函 数 名：QueryLib
'**输    入：LibInfo()(String) -
'**        ：QueryStr(String)  -
'**输    出：(Integer) -
'**功能描述：查询封装库
'**全局变量：
'**调用模块：
'**作    者：Shenhao
'**日    期：2011-10-31 23:43:41
'**修 改 人：
'**日    期：
'**版    本：V3.6.3
'*************************************************************************
Function QueryLib(LibInfo() As String, QueryStr As String) As Integer
    Dim i As Integer
    QueryLib = 0
    For i = 0 To UBound(LibInfo) - 1
        If QueryStr <> "" And LibInfo(i) <> "" Then
            If QueryStr = LibInfo(i) Then
                QueryLib = 1
                Exit For
            End If
        End If
    Next i

End Function

