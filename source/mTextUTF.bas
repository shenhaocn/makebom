Attribute VB_Name = "mTextUTF"
'*************************************************************************************
'**模 块 名：mTextUTF
'**说    明：TP-LINK SMB Switch Product Line Hardware Group 版权所有2011 - 2012(C)
'**创 建 人：zyl910(2006-1-23)
'**日    期：2011-10-31 23:38:21
'**修 改 人：Shenhao
'**日    期：
'**描    述：支持Unicode编码的文本文件读写。
'**        ：支持ANSI、UTF-8、UTF-16LE、UTF-16BE这几种编码文本
'**        ：支持UTF-8无BOM格式编码文本
'**版    本：V3.6.3
'*************************************************************************************

Option Explicit

'## 编译预处理常数 #########################################
'== 全局常数 ===============================================
'IncludeAPILib：引用了API库，此时不需要手动写API声明

'## API ####################################################
#If IncludeAPILib = 0 Then
'== File ===================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const Create_NEW = 1
Private Const Create_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

'== Unicode ================================================

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

#End If

'###########################################################

'Unicode编码格式
Public Enum UnicodeEncodeFormat
UEF_ANSI = 0 'ANSI+DBCS
UEF_UTF8     'UTF-8
UEF_UTF8NB   'UTF-8 No BOM
UEF_UTF16LE  'UTF-16LE
UEF_UTF16BE  'UTF-16BE
UEF_UTF32LE  'UTF-32LE
UEF_UTF32BE  'UTF-32BE

UEF_AUTO = -1 '自动识别编码

'隐藏项目
[_UEF_Min] = UEF_ANSI
[_UEF_Max] = UEF_UTF32BE

End Enum

'ANSI+DBCS方式的文本所使用的代码页。默认为0，表示使用系统当前代码页。
'可以利用该参数实现读取其他代码编码的文本，比如想在 简体中文平台下 读取 繁体中文平台生成的txt，就将它设为950
Public UEFCodePage As Long


'把BYTE类型变量左移1位的函数
'返回值：移位结果
'Byt：待移位的字节
Private Function ShLB_By1Bit(ByVal Byt As Byte) As Byte
    
    '(Byt And &H7F)的作用是屏蔽最高位。 *2：左移一位
    ShLB_By1Bit = (Byt And &H7F) * 2

End Function

'判断BOM
'返回值：BOM所占字节
'dwFirst：[in]文件最开始的4个字节
'fmt：[out]返回编码类型
Public Function UEFCheckBOM(ByVal dwFirst As Long, ByRef fmt As UnicodeEncodeFormat) As Long
    If dwFirst = &HFEFF& Then
        fmt = UEF_UTF32LE
        UEFCheckBOM = 4
    ElseIf dwFirst = &HFFFE0000 Then
        fmt = UEF_UTF32BE
        UEFCheckBOM = 4
    ElseIf (dwFirst And &HFFFF&) = &HFEFF& Then
        fmt = UEF_UTF16LE
        UEFCheckBOM = 2
    ElseIf (dwFirst And &HFFFF&) = &HFFFE& Then
        fmt = UEF_UTF16BE
        UEFCheckBOM = 2
    ElseIf (dwFirst And &HFFFFFF) = &HBFBBEF Then
        fmt = UEF_UTF8
        UEFCheckBOM = 3
    Else '先暂定为UEF_ANSI 后续会再区分UEF_ANSI 和 UEF_UTF8NB
        fmt = UEF_ANSI
        UEFCheckBOM = 0
    End If
End Function

'==========================================================================================
'UTF-8是一种变长的编码方式。
'它可以使用1~4个字节表示一个符号，根据不同的符号而变化字节长度。
'�
'UTF-8的编码规则很简单：
'�
'1）对于单字节的符号，字节的第一位设为0，后面7位为这个符号的unicode码。
'   因此对于英语字母，UTF-8编码和ASCII码是相同的。
'�
'2）对于n字节的符号（n>1），第一个字节的前n位都设为1，第n+1位设为0，后面字节的前两位一律设为10。
'   剩下的没有提及的二进制位，全部为这个符号的unicode码。
'�
'下表总结了编码规则，字母x表示可用编码的位。
'Unicode符号范围     | UTF-8编码方式
'(十六进制)          | （二进制）
'--------------------+---------------------------------------------
'0000 0000-0000 007F | 0xxxxxxx
'0000 0080-0000 07FF | 110xxxxx 10xxxxxx
'0000 0800-0000 FFFF | 1110xxxx 10xxxxxx 10xxxxxx
'0001 0000-0010 FFFF | 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
'==========================================================================================��


'区分UEF_ANSI 和 UEF_UTF8NB
'�
'共同点:二者均无BOM 因此调用前需要将文件指针重置到最开始的位置
'�
'bufAll：[in]文件所有字节
'fmt：[out]返回编码类型
Public Function UEFCheckUTF8NoBom(ByRef bufAll() As Byte, ByRef fmt As UnicodeEncodeFormat)
    
    Dim i As Long               '可能会溢出
    Dim cOctets As Long         '可以容纳UTF-8编码字符的字节大小 4bytes
    Dim bAllAscii As Boolean    '如果全部为ASCII，说明不是UTF-8
    
    bAllAscii = True
    cOctets = 0
    
    'Debug.Print Hex(bufAll(0)) & "-" & Hex(bufAll(1)) & "-" & Hex(bufAll(2))
    
    For i = 0 To UBound(bufAll)
        If (bufAll(i) And &H80) <> 0 Then
            'ASCII用7位储存，最高位为0，如果这里相与非0，就不是ASCII
            '对于单字节的符号，字节的第一位设为0，后面7位为这个符号的unicode码。
            '因此对于英语字母，UTF-8编码和ASCII码是相同的
            bAllAscii = False
        End If
        
        '对于n字节的符号（n>1），第一个字节的前n位都设为1，第n+1位设为0，后面字节的前两位一律设为10
        'cOctets = 0 表示本字节是leading byte
        If cOctets = 0 Then
            If bufAll(i) >= &H80 Then
                '计数：是cOctets字节的符号
                Do While (bufAll(i) And &H80) <> 0
                    'bufAll(i)左移一位
                    bufAll(i) = ShLB_By1Bit(bufAll(i))
                    cOctets = cOctets + 1
                Loop
                
                'leading byte至少应为110x xxxx
                cOctets = cOctets - 1
                If cOctets = 0 Then
                    '返回默认编码
                    fmt = UEF_ANSI
                    Exit Function
                End If
            End If
        Else
            '非leading byte形式必须是 10xxxxxx
            If (bufAll(i) And &HC0) <> &H80 Then
                '返回默认编码
                fmt = UEF_ANSI
                Exit Function
            End If
            '准备下一个byte
            cOctets = cOctets - 1
        End If
    
    Next i
    
    '文本结束.  不应有任何多余的byte 有即为错误 返回默认编码
    If cOctets > 0 Then
        fmt = UEF_ANSI
        Exit Function
    End If
    
    '如果全是ascii.  需要注意的是使用相应的code pages做转换
    If bAllAscii = True Then
        fmt = UEF_ANSI
        Exit Function
    End If
    
    '修成正果 终于格式全部正确 返回UTF8 No BOM编码格式
    fmt = UEF_UTF8NB
    
End Function

'生成BOM
'返回值：BOM所占字节
'fmt：[in]编码类型
'dwFirst：[out]文件最开始的4个字节
Public Function UEFMakeBOM(ByVal fmt As UnicodeEncodeFormat, ByRef dwFirst As Long) As Long
    Select Case fmt
    Case UEF_UTF8
        dwFirst = &HBFBBEF
        UEFMakeBOM = 3
    Case UEF_UTF16LE
        dwFirst = &HFEFF&
        UEFMakeBOM = 2
    Case UEF_UTF16BE
        dwFirst = &HFFFE&
        UEFMakeBOM = 2
    Case UEF_UTF32LE
        dwFirst = &HFEFF&
        UEFMakeBOM = 4
    Case UEF_UTF32BE
        dwFirst = &HFFFE0000
        UEFMakeBOM = 4
    Case Else 'UEF_UTF8NB和UEF_ANSI
        dwFirst = 0
        UEFMakeBOM = 0
    End Select
End Function

'判断文本文件的编码类型
'返回值：编码类型。文件无法打开时，返回UEF_Auto
'FileName：文件名
Public Function UEFCheckTextFileFormat(ByVal FileName As String) As UnicodeEncodeFormat
    Dim hFile     As Long
    Dim dwFirst   As Long
    Dim nNumRead  As Long
    
    Dim nFileSize As Long
    Dim bufAll()  As Byte

    '打开文件
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then '文件无法打开
        UEFCheckTextFileFormat = UEF_AUTO
        Exit Function
    End If

    '判断BOM
    dwFirst = 0
    Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
    nNumRead = UEFCheckBOM(dwFirst, UEFCheckTextFileFormat) '返回BOM所占字节
    'Debug.Print nNumRead
    
    '如果是判断结果是UEF_ANSI 则需继续区分UEF_ANSI 和 UEF_UTF8NB
    If UEFCheckTextFileFormat = UEF_ANSI Then
        nFileSize = GetFileSize(hFile, nNumRead)
        ReDim bufAll(0 To nFileSize - 1)
        
        nNumRead = 0
        'UEF_ANSI UEF_UTF8NB 的cbBOM均为0
        Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '恢复文件指针
        Call ReadFile(hFile, bufAll(0), nFileSize, nNumRead, ByVal 0&)
        UEFCheckUTF8NoBom bufAll, UEFCheckTextFileFormat
        
    End If

    'Debug.Print UEFCheckTextFileFormat
    
    '关闭文件
    Call CloseHandle(hFile)

End Function

'读取文本文件
'返回值：读取的文本。返回vbNullString表示文件无法打开
'FileName：[in]文件名
'fmt：[in,out]使用何种文本编码格式来读取文本。为UEF_Auto时表示自动判断，且在fmt参数返回文本所用编码格式
Public Function UEFLoadTextFile(ByVal FileName As String, Optional ByRef fmt As UnicodeEncodeFormat = UEF_AUTO) As String
    Dim hFile As Long
    Dim nFileSize As Long
    Dim nNumRead As Long
    Dim dwFirst As Long
    Dim CurFmt As UnicodeEncodeFormat
    Dim cbBOM As Long
    Dim cbTextData As Long
    Dim CurCP As Long
    Dim byBuf() As Byte
    Dim byBufDiff() As Byte
    Dim cchStr As Long
    Dim i As Long
    Dim byTemp As Byte
    
    '判断fmt范围
    If fmt <> UEF_AUTO Then
        If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
            GoTo FunEnd
        End If
    End If
    
    '打开文件
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then '文件无法打开
        GoTo FunEnd
    End If
    
    '判断文件大小
    nFileSize = GetFileSize(hFile, nNumRead)
    If nNumRead <> 0 Then '超过4GB
        GoTo FreeHandle
    End If
    If nFileSize < 0 Then '超过2GB
        GoTo FreeHandle
    End If
    
    '判断BOM
    dwFirst = 0
    Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
    cbBOM = UEFCheckBOM(dwFirst, CurFmt)
    '继续区分UEF_ANSI 和 UEF_UTF8NB cbBOM二者相同 by shenhao
    If CurFmt = UEF_ANSI Then
        ReDim byBufDiff(0 To nFileSize - 1)
        'UEF_ANSI UEF_UTF8NB 的cbBOM均为0
        Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '恢复文件指针
        Call ReadFile(hFile, byBufDiff(0), nFileSize, nNumRead, ByVal 0&)
        UEFCheckUTF8NoBom byBufDiff, CurFmt
    End If
    
    
    '恢复文件指针
    If fmt = UEF_AUTO Then '自动判断
        fmt = CurFmt
        'cbBOM = cbBOM
    Else '手动设置编码
        If fmt = CurFmt Then '若编码相同，则忽略BOM标记
            'cbBOM = cbBOM
        Else '编码不同，那么都是数据
            cbBOM = 0
        End If
    End If
    Call SetFilePointer(hFile, cbBOM, ByVal 0&, FILE_BEGIN)
    cbTextData = nFileSize - cbBOM
    
    '读取数据
    UEFLoadTextFile = ""
    Select Case fmt
        Case UEF_ANSI, UEF_UTF8, UEF_UTF8NB
            '判断应使用的CodePage
            CurCP = IIf((fmt = UEF_UTF8) Or (fmt = UEF_UTF8NB), CP_UTF8, UEFCodePage)
            
            '分配缓冲区
            On Error GoTo FreeHandle
            ReDim byBuf(0 To cbTextData - 1)
            On Error GoTo 0
            
            '读取数据
            nNumRead = 0
            Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)
            
            '取得Unicode文本长度
            cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal 0&, ByVal 0&)
            If cchStr > 0 Then
                '分配字符串空间
                On Error GoTo FreeHandle
                UEFLoadTextFile = String$(cchStr, 0)
                On Error GoTo 0
                
                '取得文本
                cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal StrPtr(UEFLoadTextFile), cchStr + 1)
            End If
            
        Case UEF_UTF16LE
            cchStr = (cbTextData + 1) / 2
            
            '分配字符串空间
            On Error GoTo FreeHandle
            UEFLoadTextFile = String$(cchStr, 0)
            On Error GoTo 0
            
            '取得文本
            nNumRead = 0
            Call ReadFile(hFile, ByVal StrPtr(UEFLoadTextFile), cbTextData, nNumRead, ByVal 0&)
            
            '修正文本长度
            cchStr = (nNumRead + 1) / 2
            If cchStr > 0 Then
                If Len(UEFLoadTextFile) > cchStr Then
                    UEFLoadTextFile = Left$(UEFLoadTextFile, cchStr)
                End If
            Else
                UEFLoadTextFile = ""
            End If
            
        Case UEF_UTF16BE
            '分配缓冲区
            On Error GoTo FreeHandle
            ReDim byBuf(0 To cbTextData - 1)
            On Error GoTo 0
            
            '读取数据
            nNumRead = 0
            Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)
            
            If nNumRead > 0 Then
                '隔两字节翻转相邻字节
                For i = 0 To nNumRead - 1 - 1 Step 2 '再-1是为了避免最后多出的那个字节
                    byTemp = byBuf(i)
                    byBuf(i) = byBuf(i + 1)
                    byBuf(i + 1) = byTemp
                Next i
                
                '取得文本
                UEFLoadTextFile = byBuf 'VB允许String中的字符串数据与Byte数组直接转换
            End If
            
        Case UEF_UTF32LE
            UEFLoadTextFile = vbNullString '暂时不支持
        Case UEF_UTF32BE
            UEFLoadTextFile = vbNullString '暂时不支持
        Case Else
            Debug.Assert False
    End Select
    
FreeHandle:
    '关闭文件
    Call CloseHandle(hFile)
    
FunEnd:

End Function

'保存文本文件
'返回值：是否成功
'FileName：[in]文件名
'sText：[in]欲输出的文本
'IsAppend：[in]是否是添加方式
'fmt：[in,out]使用何种文本编码格式来存储文本。当IsAppend=True时允许UEF_Auto自动判断，且在fmt参数返回文本所用编码格式
'DefFmt：[in]当使用添加模式时，若文件不存在且fmt = UEF_Auto时应使用的编码格式
Public Function UEFSaveTextFile(ByVal FileName As String, _
                                ByRef sText As String, _
                                Optional ByVal IsAppend As Boolean = False, _
                                Optional ByRef fmt As UnicodeEncodeFormat = UEF_AUTO, _
                                Optional ByVal DefFmt As UnicodeEncodeFormat = UEF_ANSI) As Boolean
                                
    Dim hFile As Long
    Dim nFileSize As Long
    Dim nNumRead As Long
    Dim dwFirst As Long
    Dim cbBOM As Long
    Dim CurCP As Long
    Dim byBuf() As Byte
    Dim byBufDiff() As Byte
    Dim cbBuf As Long
    Dim i As Long
    Dim byTemp As Byte
    
    '判断fmt范围
    If IsAppend And (fmt = UEF_AUTO) Then
    Else
        If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
            GoTo FunEnd
        End If
    End If
    
    '打开文件
    hFile = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, IIf(IsAppend, OPEN_ALWAYS, Create_ALWAYS), FILE_ATTRIBUTE_NORMAL, ByVal 0&)
    If INVALID_HANDLE_VALUE = hFile Then '文件无法打开
            GoTo FunEnd
    End If
    
    '判断文件大小
    nFileSize = GetFileSize(hFile, nNumRead)
    If nFileSize = 0 And nNumRead = 0 Then '文件大小为0字节
         IsAppend = False '此时需要写BOM标志
    End If
    If fmt = UEF_AUTO Then
        fmt = DefFmt
    End If
    
    '判断BOM
    If IsAppend And (fmt = UEF_AUTO) Then
        dwFirst = 0
        Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
        cbBOM = UEFCheckBOM(dwFirst, fmt)
        '继续区分UEF_ANSI 和 UEF_UTF8NB cbBOM二者相同 by shenhao
        If fmt = UEF_ANSI Then
            ReDim byBufDiff(0 To nFileSize - 1)
            'UEF_ANSI UEF_UTF8NB 的cbBOM均为0
            Call SetFilePointer(hFile, 0, ByVal 0&, FILE_BEGIN) '恢复文件指针
            Call ReadFile(hFile, byBufDiff(0), nFileSize, nNumRead, ByVal 0&)
            UEFCheckUTF8NoBom byBufDiff, fmt
        End If
        
    ElseIf IsAppend = False Then
        cbBOM = UEFMakeBOM(fmt, dwFirst)
    End If
    
    '文件指针定位
    Call SetFilePointer(hFile, 0, ByVal 0&, IIf(IsAppend, FILE_END, FILE_BEGIN))
    
    '写BOM
    If IsAppend = False Then
        If cbBOM > 0 Then
            Call WriteFile(hFile, dwFirst, cbBOM, nNumRead, ByVal 0&)
        End If
    End If
    
    '写文本数据
    If Len(sText) > 0 Then
        Select Case fmt
            Case UEF_ANSI, UEF_UTF8, UEF_UTF8NB
                '判断应使用的CodePage
                CurCP = IIf((fmt = UEF_UTF8) Or (fmt = UEF_UTF8NB), CP_UTF8, UEFCodePage)
                
                '取得缓冲区大小
                cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), ByVal 0&, 0, ByVal 0&, ByVal 0&)
                If cbBuf > 0 Then
                    '分配缓冲区
                    On Error GoTo FreeHandle
                    ReDim byBuf(0 To cbBuf)
                    On Error GoTo 0
                
                    '转换文本
                    cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), byBuf(0), cbBuf + 1, ByVal 0&, ByVal 0&)
                
                    '写文件
                    Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)
                
                    UEFSaveTextFile = True
                End If
                
            Case UEF_UTF16LE
                '写文件
                Call WriteFile(hFile, ByVal StrPtr(sText), LenB(sText), nNumRead, ByVal 0&)
            
                UEFSaveTextFile = True
            
            Case UEF_UTF16BE
                '将字符串中的数据复制到byBuf
                On Error GoTo FreeHandle
                byBuf = sText
                On Error GoTo 0
                cbBuf = UBound(byBuf) - LBound(byBuf) + 1
            
                '隔两字节翻转相邻字节
                For i = 0 To cbBuf - 1 - 1 Step 2 '再-1是为了避免最后多出的那个字节
                    byTemp = byBuf(i)
                    byBuf(i) = byBuf(i + 1)
                    byBuf(i + 1) = byTemp
                Next i
            
                '写文件
                Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)
            
                UEFSaveTextFile = True
            
            Case UEF_UTF32LE
                UEFSaveTextFile = False '暂时不支持
            Case UEF_UTF32BE
                UEFSaveTextFile = False '暂时不支持
            Case Else
                Debug.Assert False
        End Select
    Else
        UEFSaveTextFile = True
    End If
    
FreeHandle:
    '关闭文件
    Call CloseHandle(hFile)
    
FunEnd:
End Function
