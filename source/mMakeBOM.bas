Attribute VB_Name = "mMakeBOM"
'*************************************************************************************
'**Ä£ ¿é Ãû£ºmMakeBOM
'**Ëµ    Ã÷£ºTP-LINK SMB Switch Product Line Hardware Group °æÈ¨ËùÓĞ2011 - 2012(C)
'**´´ ½¨ ÈË£ºShenhao
'**ÈÕ    ÆÚ£º2011-10-31 23:37:45
'**ĞŞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**Ãè    Êö£ºExcel¸ñÊ½BOMÉú³É
'**°æ    ±¾£ºV3.6.3
'*************************************************************************************
Option Explicit

'BOMÀàĞÍ
Public Enum BomType

BOM_ALL = 0
BOM_NCDBG
BOM_NONE

BOM_Ô¤
BOM_ÁìÁÏ
BOM_µ÷ÊÔ
BOM_Éú²ú

End Enum

'BMFÎÄ¼ş±àÂë¸ñÊ½
'Item Number Part Number Value   Quantity    Part Reference  PCB Footprint Mount Type Description TP1 TP2 TP3
'0-----------1-----------2-------3-----------4---------------5-------------6----------7-----------8---9--10--
'BMF_ItemNum=0
'BMF_PartNum
'BMF_Value
'BMF_Quantity
'BMF_PartRef
'BMF_PcbFB
'BMF_MountType
'BMF_Description
'BMF_TP1
'BMF_TP2
'BMF_TP3ÿ


'*************************************************************************
'**º¯ Êı Ãû£ºxlsInsert
'**Êä    Èë£ºxlSheet(Excel.Worksheet) -
'**        £ºItemNum(Integer)         -
'**        £ºRow(Long)                -
'**        £ºinsertStr()(String)      -
'**        £ºOrgEnable(Boolean)       -
'**Êä    ³ö£ºÎŞ
'**¹¦ÄÜÃèÊö£ºÔÚ¶ÔÓ¦µÄSheetµÄ¶ÔÓ¦Î»ÖÃÉÏÌí¼ÓÒ»ĞĞ
'**È«¾Ö±äÁ¿£º
'**µ÷ÓÃÄ£¿é£º
'**×÷    Õß£ºShenhao
'**ÈÕ    ÆÚ£º2011-11-01 00:37:31
'**ĞŞ ¸Ä ÈË£º
'**ÈÕ    ÆÚ£º
'**°æ    ±¾£ºV3.6.7
'*************************************************************************
Function xlsInsert(xlSheet As Excel.Worksheet, ItemNum As Integer, Row As Long, insertStr() As String, OrgEnable As Boolean)
    
    'Ê×ĞĞ²»ĞèÒª¼ÓÈë
    If ItemNum > 1 Then
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Insert
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Interior.Pattern = xlNone 'È¥³ıÑÕÉ«µÈ¸ñÊ½ ĞŞÕıÏÔÊ¾bug
    End If
    
    xlSheet.Cells(ItemNum + Row, 1) = ItemNum
    xlSheet.Cells(ItemNum + Row, 2) = insertStr(BMF_PartNum)
    xlSheet.Cells(ItemNum + Row, 3) = insertStr(BMF_Description)
    xlSheet.Cells(ItemNum + Row, 5) = insertStr(BMF_Quantity)
    xlSheet.Cells(ItemNum + Row, 6) = insertStr(BMF_PartRef)
    xlSheet.Cells(ItemNum + Row, 7) = insertStr(BMF_PcbFB)
    xlSheet.Cells(ItemNum + Row, 8) = insertStr(BMF_Value)
    
    'ÊÇ·ñÌí¼Ó¿â´æĞÅÏ¢£¿
    If OrgEnable = True Then
    
        'Ìí¼ÓTP1¿â´æĞÅÏ¢
        If insertStr(BMF_TP1) = "-" Then
            xlSheet.Cells(ItemNum + Row, 9) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 9) = insertStr(BMF_TP1)
            If insertStr(BMF_TP1) = "0" Or InStr(insertStr(BMF_TP1), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 9).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
        
        'Ìí¼ÓTP2¿â´æĞÅÏ¢
        If insertStr(BMF_TP2) = "-" Then
            xlSheet.Cells(ItemNum + Row, 10) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 10) = insertStr(BMF_TP2)
            If insertStr(BMF_TP2) = "0" Or InStr(insertStr(BMF_TP2), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 10).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
            
        'Ìí¼ÓTP3¿â´æĞÅÏ¢
        If insertStr(BMF_TP3) = "-" Then
            xlSheet.Cells(ItemNum + Row, 11) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 11) = insertStr(BMF_TP3)
            If insertStr(BMF_TP3) = "0" Or InStr(insertStr(BMF_TP3), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 11).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
        
    End If
    
End Function

'ÅÅĞòÊı×é£¨¿ÉÒÔÊÇ×Ö·û´®Êı×é£©£ºBigToSmall=True ´Ó´óµ½Ğ¡£¬·ñÔò ´ÓĞ¡´óµ½
Function ZuSorted(Zu() As Variant, prefixStr As String, Optional BigToSmall As Boolean) As String
   
   Dim i As Long, j As Long, S As Variant
   Dim TF As Boolean, nL As Long, nU As Long
   
   nL = LBound(Zu): nU = UBound(Zu)
   For i = nL To nU
      For j = nL To nU '       ´Ó´óµ½Ğ¡                ´ÓĞ¡´óµ½
        If BigToSmall Then TF = Zu(j) < Zu(i) Else TF = Zu(j) > Zu(i)
        If TF Then S = Zu(i): Zu(i) = Zu(j): Zu(j) = S
      Next j
   Next i
   
   For i = nL To nU - 1
      ZuSorted = ZuSorted + prefixStr + CStr(Zu(i)) + Space(1)
   Next i
   
   ZuSorted = ZuSorted + prefixStr + CStr(Zu(nU))
   
End Function

'ÓÉÓÚÎ»ºÅ³¤¶È²»Ò»£¬³£¹æÅÅĞò·½·¨²»¿ÉĞĞ Òò´ËĞèÒªÌØÊâÅÅĞò·½·¨
Function RealSorted(ByRef RefStr As String, Optional BigToSmall As Boolean) As Boolean
    Dim srcStr()    As String
    Dim intSorted() As Variant
    Dim i           As Long
    Dim Index       As Long 'Î»ºÅºóÊı×Ö¿ªÊ¼µÄÎ»ÖÃ
    
    Dim prefixStr   As String
    
    RealSorted = False
    
    srcStr = Split(RefStr, Space(1))
    ReDim intSorted(UBound(srcStr))
    
    Index = 0
    For i = 0 To Len(srcStr(0))
        If IsNumeric(Right(srcStr(0), Len(srcStr(0)) - i)) = True Then
            prefixStr = Left(srcStr(0), i)
            Index = i
            Exit For
        End If
    Next
    
    For i = LBound(intSorted) To UBound(intSorted)
        If IsNumeric(Right(srcStr(i), Len(srcStr(i)) - Index)) = True Then
            intSorted(i) = Val(Right(srcStr(i), Len(srcStr(i)) - Index))
        Else
            RealSorted = False
        End If
    Next i
    
    RefStr = ZuSorted(intSorted, prefixStr, BigToSmall)
    RealSorted = True
    
End Function

'Ìí¼ÓDBG NCÔª¼şµ½ÏàÓ¦µÄExcel BOMÖĞ ¶ÔºÏ²¢µÄÎ»ºÅ½øĞĞÅÅĞò
Function addDbgNcPart(xlSheet As Excel.Worksheet, bmfAtom() As String, _
                      ByRef ItemNum1 As Integer, ByRef ItemNum2 As Integer, _
                      rngPos1 As Range, rngPos2 As Range, OrgEnable As Boolean)

    Dim rngNum       As Range
    Dim partRefStr() As String
    Dim tmpRefStr    As String
    Dim i            As Integer
    
    'ÊÇ·ñĞèÒªºÏ²¢Ôª¼ş
    With xlSheet.Cells
        Set rngNum = .Find(bmfAtom(BMF_PartNum), lookin:=xlValues)
        If rngNum Is Nothing Then
            '²»ĞèÒªºÏ²¢ Ö±½ÓÌí¼ÓÔÚºóÃæ
            Select Case bmfAtom(BMF_MountType)
                Case "S"
                ItemNum1 = ItemNum1 + 1
                xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Font.ColorIndex = 5

                Case "S+"
                ItemNum1 = ItemNum1 + 1
                xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Interior.Color = 16737792
                xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Font.ColorIndex = 5

                Case "L"
                ItemNum2 = ItemNum2 + 1
                xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
                xlSheet.Rows((rngPos2.Row + ItemNum2) & ":" & (rngPos2.Row + ItemNum2)).Font.ColorIndex = 5

                Case "N"
                'Do Nothing

                Case Else
                '¿âÎÄ¼şÖĞÃ»ÓĞ²éµ½·â×°£¬¾Ü¾øÉú³ÉBOM
                MsgBox "Î´Öª·â×°[" & bmfAtom(BMF_PcbFB) & "]£¬Çë¸üĞÂ¿âÎÄ¼ş£¡"
                
            End Select
            
        Else 'ĞèÒªºÏ²¢ ĞŞ¸ÄÊıÁ¿(µÚ5ÁĞ) ºÍÎ»ºÅ(µÚ6ÁĞ) Î»ºÅĞèÒªÖØĞÂÅÅĞò
            'xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            xlSheet.Cells(rngNum.Row, 5) = CInt(xlSheet.Cells(rngNum.Row, 5)) + CInt(bmfAtom(BMF_Quantity))
            xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
            
            'Î»ºÅĞèÒªÅÅĞò£¡
            tmpRefStr = xlSheet.Cells(rngNum.Row, 6) + " " + bmfAtom(BMF_PartRef)
            RealSorted tmpRefStr, False
            xlSheet.Cells(rngNum.Row, 6) = tmpRefStr
            xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
        End If
    End With
End Function

'ĞĞ¿½±´
Function CopyLine(xlSheetTo As Excel.Worksheet, RowTo As Integer, xlSheetFrom As Excel.Worksheet, RowFrom As Integer, ColumnNum As Integer, PartNum As Integer)
    xlSheetTo.Rows(RowTo & ":" & RowTo).Insert
    xlSheetTo.Cells(RowTo, 1) = PartNum
    Dim i As Integer
    For i = 2 To ColumnNum
        xlSheetTo.Cells(RowTo, i) = xlSheetFrom.Cells(RowFrom, i)
    Next i
    xlSheetTo.Rows(RowTo & ":" & RowTo).Font.ColorIndex = 5
End Function

'¸ù¾İÑ¡Ïî´´½¨BOM ²¢µ÷Õû¸ñÊ½ ÎªÌî³äÊı¾İ×ö×¼±¸
Function ExcelCreate(bt_value As BomType)

    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    '=====================================================================================
    'Ô¤BOM£ºCaptureÖĞµ¼³öµÄBOM³ıNoneÔª¼ş¡¢NCÔª¼ş¡¢DBGÔª¼ş¡¢DBG_NCÔª¼şÖ®ÍâµÄËùÓĞÔª¼şµÄ¼¯ºÏ¡£
    '=====================================================================================
    If bt_value = BOM_Ô¤ Then
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_Ô¤BOM_BMF.xls")
        
        xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    End If
    
    '=====================================================================================
    'NC_DBGÔª¼şxls
    '=====================================================================================
    If bt_value = BOM_NCDBG Then
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_NC_DBG.xls")
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'NoneÔª¼şxls
    '=====================================================================================
    If bt_value = BOM_NONE Then
    
        Dim rngNC           As Range
        Dim rngDB           As Range
        Dim rngDBNC         As Range
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        xlBook.SaveAs (SaveAsPath & "_None_PartRef.xls")
        xlBook.Worksheets(1).Name = "NoneÔª¼ş"
        
        Set xlSheet = xlBook.Worksheets(1)
        
        With xlSheet.Cells
            Set rngNC = .Find("NCÔª¼ş", lookin:=xlValues)
            Set rngDB = .Find("DBGÔª¼ş", lookin:=xlValues)
            Set rngDBNC = .Find("DBG_NCÔª¼ş", lookin:=xlValues)
            If rngNC Is Nothing Or rngDB Is Nothing Then
                MsgBox "NC_DBGÄ£°å´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
                End
            End If
        End With
        
        'µ÷ÕûNoneShleet
        xlSheet.Cells(rngNC.Row, 2) = "None"
        xlSheet.Rows(rngDB.Row & ":" & rngDBNC.Row + 1).Delete
    
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼ş - ĞÂ´òÑùÎïÁÏ+ ÎïÁÏ¿â´æĞÅÏ¢¡£
    '=====================================================================================
    If bt_value = BOM_ÁìÁÏ Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_ÁìÁÏBOM.xls")
         
        'µ÷ÕûÁìÁÏBOM¸ñÊ½
        'ÔÚÁìÁÏBOMÖĞ²åÈëÁĞ(I£ºTP1¿â´æ) (J£ºTP2¿â´æ) (K£ºTP3¿â´æ)£¨ĞèÑ¡Ôñ¿â´æĞÅÏ¢£©
        'µ÷Õû±í¸ñÊÊÓ¦ĞÂÌí¼ÓµÄÄÚÈİ
        xlSheet.Columns("C:C").ColumnWidth = 45
        xlSheet.Columns("G:G").ColumnWidth = 12
        xlSheet.Columns("H:H").ColumnWidth = 12
        
        xlSheet.Columns("H:H").Copy
        xlSheet.Columns("I:I").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
        xlSheet.Columns("I:I").Copy
        xlSheet.Columns("J:J").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
        xlSheet.Columns("K:K").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
        xlApp.CutCopyMode = False
        
        'Ìî³äÊı¾İ µÚ5ĞĞ²ÅÊÇÕæÕıµÄÁĞÍ·ÃèÊöĞĞ
        xlSheet.Cells(5, 9) = "TP1¿â´æ"
        xlSheet.Cells(5, 10) = "TP2¿â´æ"
        xlSheet.Cells(5, 11) = "TP3¿â´æ"

        xlSheet.Cells(5, 1).Select
         
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼ş ¼¯ºÏ¡£
    '=====================================================================================
    If bt_value = BOM_µ÷ÊÔ Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlSheet.SaveAs (SaveAsPath & "_µ÷ÊÔBOM.xls")
        
        'µ÷Õû´òÓ¡¸ñÊ½
        With xlSheet.PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .Zoom = 80
        End With
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼ş ¼¯ºÏ
    '=====================================================================================
    If bt_value = BOM_Éú²ú Then

        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_Éú²úBOM.xls")
        
        xlBook.Close (True)
    End If
    
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    Exit Function
    
ErrorHandle:

    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    
    MsgBox "´´½¨BOMÖĞ¼äÎÄ¼şÊ±·¢ÉúÒì³£", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
    

End Function

'Ìî³äÏàÓ¦µÄÊı¾İ

'=====================================================================================
'1.Ô¤BOM£ºCaptureÖĞµ¼³öµÄBOM³ıNoneÔª¼ş¡¢NCÔª¼ş¡¢DBGÔª¼ş¡¢DBG_NCÔª¼şÖ®ÍâµÄËùÓĞÔª¼şµÄ¼¯ºÏ¡£
'2.NC_DBGÔª¼şxls
'3.NoneÔª¼şxls
'4.ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼ş - ĞÂ´òÑùÎïÁÏ+ ÎïÁÏ¿â´æĞÅÏ¢
'5.µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼ş ¼¯ºÏ
'6.Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼ş ¼¯ºÏ
'ÿ
'Ôª¼şÀàĞÍ £º ÆÕÍ¨Ôª¼ş NcDbgÔª¼ş NoneÔª¼ş ´òÑùÎïÁÏ ÿ
'=====================================================================================
Function CreateBOM(bt_value As BomType) As Boolean

    On Error GoTo ErrorHandle

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    '´ò¿ªÏàÓ¦µÄÎÄ¼ş
    Select Case bt_value
        Case BOM_NCDBG

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
        Case BOM_NONE

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
        Case BOM_Ô¤

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_Ô¤BOM_BMF.xls")
        Case BOM_ÁìÁÏ

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_ÁìÁÏBOM.xls")
        Case BOM_µ÷ÊÔ

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_µ÷ÊÔBOM.xls")
        Case BOM_Éú²ú

        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_Éú²úBOM.xls")
        Case Else
        GoTo ErrorHandle
    End Select

    Set xlSheet = xlBook.Worksheets(1)

    '¶¨Î»¸÷ÖÖÔª¼şÎ»ÖÃ
    Dim rngPos1       As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"SMTÔª¼ş"Î»ÖÃ ÔÚDBG¡¢NoneÔª¼ş´ú±í"NCÔª¼ş"Î»ÖÃ
    Dim rngPos2       As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"DIPÔª¼ş"Î»ÖÃ ÔÚDBG¡¢NoneÔª¼ş´ú±í"DBGÔª¼ş"Î»ÖÃ
    Dim rngPos3       As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"ÆäËûÔª¼ş"Î»ÖÃ ÔÚDBG¡¢NoneÔª¼ş´ú±í"DBG_NCÔª¼ş"Î»ÖÃ
    With xlSheet.Cells

        Select Case bt_value
            Case BOM_NCDBG

            Set rngPos1 = .Find("NCÔª¼ş", lookin:=xlValues)
            Set rngPos2 = .Find("DBGÔª¼ş", lookin:=xlValues)
            Set rngPos3 = .Find("DBG_NCÔª¼ş", lookin:=xlValues)

            Case BOM_NONE

            Set rngPos1 = .Find("None", lookin:=xlValues)
            Set rngPos2 = .Find("None", lookin:=xlValues)
            Set rngPos3 = .Find("None", lookin:=xlValues)

            Case BOM_Ô¤, BOM_ÁìÁÏ, BOM_µ÷ÊÔ, BOM_Éú²ú

            Set rngPos1 = .Find("SMTÔª¼ş", lookin:=xlValues)
            Set rngPos2 = .Find("DIPÔª¼ş", lookin:=xlValues)
            Set rngPos3 = .Find("ÆäËûÔª¼ş", lookin:=xlValues)

            Case Else
            GoTo ErrorHandle
        End Select

        If rngPos1 Is Nothing Or rngPos2 Is Nothing Or rngPos3 Is Nothing Then
            MsgBox "Ä£°å´íÎó-Ôª¼şÎ»ÖÃ¶¨Î»´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
            GoTo ErrorHandle
        End If
    End With

    '=======================================================
    '¿ªÊ¼¶ÁÈ¡bmfÎÄ¼ş²¢Ìî³äÏàÓ¦µÄÄÚÈİ
    '=======================================================
    Dim bmfBomLine()    As String
    Dim bmfAtom()       As String

    Dim ItemNum1        As Integer 'NCÔª¼ş »ò NoneÔª¼ş »ò SMTÔª¼şÊı
    Dim ItemNum2        As Integer 'DBGÔª¼ş »ò NoneÔª¼ş »ò DIPÔª¼şÊı
    Dim ItemNum3        As Integer 'DBG_NCÔª¼ş »ò NoneÔª¼ş »ò ÆäËûÔª¼şÊı

    Dim i               As Integer
    Dim OrgEnable       As Boolean

    ItemNum1 = 0
    ItemNum2 = 0
    ItemNum3 = 0

    'ÊÇ·ñÌí¼Ó¿â´æĞÅÏ¢£¿
    'ĞèÒª¼ì²âtsvÎÄ¼ş´´½¨Ê±¼ä Èç¹ûÊ±¼ä²»ÔÚ3ÌìÄÚ ¿â´æĞÅÏ¢Ìí¼ÓÁËÒ²ÊÇÃ»ÓĞÓÃµÄ
    'ÒÑ¾­±¾Functionµ÷ÓÃÇ°¼ì²âÁËtsvÎÄ¼şÊ±¼ä
    If bt_value = BOM_ÁìÁÏ Then
        OrgEnable = True
    Else
        OrgEnable = False
    End If

    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If

    '¶ÁÈ¡bmfÎÄ¼ş ²¢°´ĞĞ·Ö¸îÎªÊı×é
    bmfBomLine = Split(GetFileContents(BmfFilePath), vbCrLf)

    '±éÀúbmfÎÄ¼ş
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)

        If bmfAtom(BMF_PcbFB) = "" Or bmfAtom(BMF_Value) = "" Then
            MsgBox "µÚ" & bmfAtom(BMF_ItemNum) & "ÏîÔª¼şĞÅÏ¢²»ÍêÕû", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "¾¯¸æ"
            GoTo ErrorHandle
        End If

        If InStr(bmfAtom(BMF_Value), "_DBG_NC") > 0 Or bmfAtom(BMF_Value) = "DBG_NC" Then
            If bt_value = BOM_NCDBG Then
                'DBG_NCÔª¼ş
                ItemNum3 = ItemNum3 + 1
                xlsInsert xlSheet, ItemNum3, rngPos3.Row, bmfAtom, OrgEnable
            End If

        ElseIf InStr(bmfAtom(BMF_Value), "_DBG") > 0 Or bmfAtom(BMF_Value) = "DBG" Then
            If bt_value = BOM_NCDBG Then
                'DBGÔª¼ş
                ItemNum2 = ItemNum2 + 1
                xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            End If

        ElseIf InStr(bmfAtom(BMF_Value), "_NC") > 0 Or bmfAtom(BMF_Value) = "NC" Then
            If bt_value = BOM_NCDBG Then
                'NCÔª¼ş
                ItemNum1 = ItemNum1 + 1
                xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
            End If

        Else
            If bt_value = BOM_NONE Then
                If bmfAtom(BMF_MountType) = "N" Then
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                End If
            End If

            If bt_value = BOM_µ÷ÊÔ Or bt_value = BOM_ÁìÁÏ Or _
               bt_value = BOM_Éú²ú Or bt_value = BOM_Ô¤ Then
                '========================================================
                'ÆÕÍ¨Ôª¼ş Çø·ÖÔª¼şÌù×°ÀàĞÍ ÏÈ²»Çø·ÖÕâËÄ¸öBOM ºóĞøµ÷Õû
                Select Case bmfAtom(BMF_MountType)
                    Case "S"
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable

                    Case "S+"
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                    xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Interior.Color = 16737792

                    Case "L"
                    ItemNum2 = ItemNum2 + 1
                    xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable

                    Case "N"
                    'Do Nothing

                    Case Else
                    '¿âÎÄ¼şÖĞÃ»ÓĞ²éµ½·â×°£¬¾Ü¾øÉú³ÉBOM
                    MsgBox "Î´Öª·â×°[" & bmfAtom(BMF_PcbFB) & "]£¬Çë¸üĞÂ¿âÎÄ¼ş£¡"
                    GoTo ErrorHandle
                    
                End Select
            End If

        End If

    Next i

    If bt_value = BOM_µ÷ÊÔ Or bt_value = BOM_ÁìÁÏ Or _
       bt_value = BOM_Éú²ú Or bt_value = BOM_Ô¤ Then
        'ĞŞ¸Ä»úĞÍÃû³Æ
        xlSheet.Cells(2, 1) = "»úĞÍ£º  " & MainForm.ItemNameText.Text & "            PCBA °æ±¾£º                       °ë³ÉÆ·±àºÅ£º"
        If MainForm.ItemNameText.Text = "" Then
            xlSheet.Cells(2, 1).Font.ColorIndex = 5
        End If

    End If

    If bt_value = BOM_µ÷ÊÔ Or bt_value = BOM_ÁìÁÏ Or bt_value = BOM_Éú²ú Then
        'µ÷ÕûÇø·Ö¸÷ÖÖ²»Í¬µÄBOM
        'Ô¤BOM£ºCaptureÖĞµ¼³öµÄBOM³ıNoneÔª¼ş¡¢NCÔª¼ş¡¢DBGÔª¼ş¡¢DBG_NCÔª¼şÖ®ÍâµÄËùÓĞÔª¼şµÄ¼¯ºÏ
        'ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼ş - ĞÂ´òÑùÎïÁÏ + ÎïÁÏ¿â´æĞÅÏ¢
        'µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼ş ¼¯ºÏ
        'Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼ş ¼¯ºÏ
        
        'ÖØĞÂ±éÀúbmfÎÄ¼ş ¸ù¾İµ÷ÕûÏàÓ¦BOMĞÅÏ¢
         For i = 1 To UBound(bmfBomLine) - 1
            bmfAtom = Split(bmfBomLine(i), vbTab)
    
            If InStr(bmfAtom(BMF_Value), "_DBG_NC") > 0 Or bmfAtom(BMF_Value) = "DBG_NC" Then
            
                If bt_value = BOM_Éú²ú Then
                    addDbgNcPart xlSheet, bmfAtom, ItemNum1, ItemNum2, rngPos1, rngPos2, OrgEnable
                End If
    
            ElseIf InStr(bmfAtom(BMF_Value), "_DBG") > 0 Or bmfAtom(BMF_Value) = "DBG" Then
            
                If bt_value = BOM_µ÷ÊÔ Or bt_value = BOM_ÁìÁÏ Then
                    addDbgNcPart xlSheet, bmfAtom, ItemNum1, ItemNum2, rngPos1, rngPos2, OrgEnable
                End If
    
            End If
    
        Next i
        
        'É¾³ı´òÑùÔª¼ş
        If bt_value = BOM_ÁìÁÏ Then
            If DelSamplePart(xlSheet) = False Then
                GoTo ErrorHandle
            End If
        End If
    
        'ÖØĞÂ±àºÅ
        If ReNum(xlSheet) = False Then
            GoTo ErrorHandle
        End If
    End If
    
    'ÊÇ·ñ¾­¹ıBomChecker£¿
    '¾­¹ıÄ£°æµÄExcelÎÄ¼ş¸ñÊ½Ó¦¸ÃÊÇOKµÄ
    
    '½áÊøËùÓĞµ÷ÕûºÍ´´½¨ Ìí¼Ó
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾

    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó

    Exit Function

ErrorHandle:
    
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó

    MsgBox "Éú³ÉBOMÖĞ¼äÎÄ¼şÊ±·¢ÉúÒì³£", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"

End Function

'É¾³ı´òÑùÎïÁÏ
Function DelSamplePart(xlSheet As Excel.Worksheet) As Boolean
    Dim rngStart        As Range
    Dim rngEND          As Range
    
    'É¾³ı´òìÈÎïÁÏ ÁÏºÅ´æÔÚ µ«ÊÇÊôÓÚ12345xxxxx »òxxxxx xxxxxÀàĞÍ
    With xlSheet.Cells
        Set rngStart = .Find("SMTÔª¼ş", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngStart Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOMÄ£°å´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
            DelSamplePart = False
        End If
    End With
    
    Dim i          As Integer
    Dim j          As Integer
    Dim PartNum    As Integer   'Î±Ôª¼ş¸öÊı
    Dim DelRows()  As Integer   '¼ÇÂ¼ÒªÉ¾³ıµÄĞĞºÅ
    
    PartNum = rngEND.Row - rngStart.Row
    ReDim DelRows(PartNum) As Integer
    
    j = 0
    For i = rngStart.Row To rngEND.Row
        If IsNumeric(xlSheet.Cells(i, 2)) = False _
             And xlSheet.Cells(i, 2) <> "SMTÔª¼ş" _
             And xlSheet.Cells(i, 2) <> "DIPÔª¼ş" _
             And xlSheet.Cells(i, 2) <> "ÆäËûÔª¼ş" _
             And xlSheet.Cells(i, 2) <> "END" Then
            DelRows(j) = i
            j = j + 1
        End If
    Next i
    
    For i = 0 To j
        If DelRows(i) <> 0 Then
            xlSheet.Rows(DelRows(i) - i & ":" & DelRows(i) - i).Delete
        End If
    Next i
    
    DelSamplePart = True
    
End Function

'ÖØĞÂ±àºÅ
Function ReNum(xlSheet As Excel.Worksheet) As Boolean
    'ÖØĞÂ±àºÅ
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    Dim rngEND          As Range
    
    'É¾³ı´òìÈÎïÁÏ ÁÏºÅ´æÔÚ µ«ÊÇÊôÓÚ12345xxxxx »òxxxxx xxxxxÀàĞÍ
    With xlSheet.Cells
        Set rngSMT = .Find("SMTÔª¼ş", lookin:=xlValues)
        Set rngLEAD = .Find("DIPÔª¼ş", lookin:=xlValues)
        Set rngOther = .Find("ÆäËûÔª¼ş", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngSMT Is Nothing Or rngLEAD Is Nothing Or rngOther Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOMÄ£°å´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
            ReNum = False
        End If
    End With
    
    Dim j As Integer
    For j = 1 To rngLEAD.Row - rngSMT.Row - 1
        xlSheet.Cells(rngSMT.Row + j, 1) = j
    Next j
    
    For j = 1 To rngOther.Row - rngLEAD.Row - 1
        xlSheet.Cells(rngLEAD.Row + j, 1) = j
    Next j
    
    For j = 1 To rngEND.Row - rngOther.Row - 1
        xlSheet.Cells(rngOther.Row + j, 1) = j
    Next j
    
    ReNum = True
    
End Function

'¼ì²éBOM
'a.µ÷ÕûÎ»ºÅ¶ÔÓ¦µÄÁĞ¿í?ĞĞ¸ßÊ¹Ö®È«²¿ÏÔÊ¾
'b.Ôª¼şÊıºÍÎ»ºÅÊıÒ»ÖÂ
'c.Ìí¼Ó±¸×¢ÏµÁĞ
'c.1.FlashÔö¼Ó±¸×¢: ĞèÔ¤±à³Ì
'c.2.ÉÕÂ¼Èí¼şÌí¼Ó±¸×¢: SMTÓÃ
'c.3.¹¤¾ßÈí¼şÌí¼Ó±¸×¢£º"²âÊÔ½×¶ÎÓÃ"
Function BomChecker(ExcelBomFilePath As String)
    On Error GoTo ErrorHandle

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    '¶¨Î»ÔªËØÎ»ÖÃ ¶¨Î»ÁĞµÄÎ»ÖÃ
    Dim rngNum    As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"ĞòºÅ"µÄÎ»ÖÃ
    Dim rngDcp    As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"¹æ¸ñĞÍºÅ"µÄÎ»ÖÃ
    Dim rngNote   As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"¸¨ÖúËµÃ÷"µÄÎ»ÖÃ
    Dim rngQty    As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"ÊıÁ¿"µÄÎ»ÖÃ
    Dim rngRef    As Range 'ÔÚBOMÎÄ¼şÖĞ±íÊ¾"Î»ºÅ"µÄÎ»ÖÃ
    
    Dim usedRow   As Integer  '×ÜĞĞÊı
    Dim usedCol   As Integer  '×ÜÁĞÊı
    
    Dim BomAtom() As String
    Dim tmpRefStr As String

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Process 10, "´ò¿ªÎÄ¼ş...."
    
    '´ò¿ªÎÄ¼ş
    Set xlBook = xlApp.Workbooks.Open(ExcelBomFilePath)
    'ÏÈ¼ì²éµÚµÚÒ»¸öWorkSheet
    Set xlSheet = xlBook.Worksheets(1)
    
    Process 15, "ÑéÖ¤ÎÄ¼ş...."
    With xlSheet.Cells

        Set rngNum = .Find("ĞòºÅ", lookin:=xlValues)
        Set rngDcp = .Find("¹æ¸ñĞÍºÅ", lookin:=xlValues)
        Set rngNote = .Find("¸¨ÖúËµÃ÷", lookin:=xlValues)
        Set rngQty = .Find("ÊıÁ¿", lookin:=xlValues)
        Set rngRef = .Find("Î»ºÅ", lookin:=xlValues)
        
    End With
    
    If rngNum Is Nothing Or rngNote Is Nothing Or rngQty Is Nothing Or rngRef Is Nothing Then
        '¼ì²éµÚµÚ¶ş¸öWorkSheet
        Set xlSheet = xlBook.Worksheets(2)
        With xlSheet.Cells
    
            Set rngNum = .Find("ĞòºÅ", lookin:=xlValues)
            Set rngDcp = .Find("¹æ¸ñĞÍºÅ", lookin:=xlValues)
            Set rngNote = .Find("¸¨ÖúËµÃ÷", lookin:=xlValues)
            Set rngQty = .Find("ÊıÁ¿", lookin:=xlValues)
            Set rngRef = .Find("Î»ºÅ", lookin:=xlValues)
        End With
        
        If rngNum Is Nothing Or rngNote Is Nothing Or rngQty Is Nothing Or rngRef Is Nothing Then
                '½öÖ§³ÖÇ°Á½¸öSheet ÆäËûµÄ²»Ö§³Ö
                MsgBox "BOMÔªËØÎ»ÖÃ¶¨Î»´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
                GoTo ErrorHandle
            End If
    End If
    
    Process 20, "ÔªËØÎ»ÖÃÍê³É...."
    
    usedRow = xlSheet.UsedRange.Rows.Count
    usedCol = xlSheet.UsedRange.Columns.Count
    '===============================================================================================
    '¿ªÊ¼Check
    '===============================================================================================
    Dim j As Integer
    For j = rngNum.Row + 1 To usedRow
    
    'a.µ÷ÕûÎ»ºÅ¶ÔÓ¦µÄÁĞ¿í?ĞĞ¸ßÊ¹Ö®È«²¿ÏÔÊ¾
        'ÏÈÕûÀíÄÚÈİ È¥³ıÈí»Ø³µ ¶àÓàµÄ¿Õ¸ñµÈµÈ
        xlSheet.Cells(j, rngRef.Column) = clearRefStr(xlSheet.Cells(j, rngRef.Column))
        
        'µ÷ÕûÎ»ºÅÅÅĞò
        Process j * 70 / (usedRow - rngNum.Row) + 20, "ÅÅĞò" & "[" & j & "]" & "ĞĞÎ»ºÅ..."
        If xlSheet.Cells(j, rngRef.Column) <> "" Then
            tmpRefStr = xlSheet.Cells(j, rngRef.Column)
            If RealSorted(tmpRefStr, False) = True Then
                xlSheet.Cells(j, rngRef.Column) = tmpRefStr
            Else
                MsgBox "µÚ[" & j & "]ĞĞÎ»ºÅ¸ñÊ½´íÎó£¬ÎŞ·¨½øĞĞÖØĞÂÅÅĞò£¡", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
                GoTo ErrorHandle
            End If
        End If
        
        'µ÷ÕûÎ»ºÅ¶ÔÓ¦µÄÁĞ¿í?ĞĞ¸ßÊ¹Ö®È«²¿ÏÔÊ¾
        Process j * 70 / (usedRow - rngNum.Row) + 20, "µ÷Õû" & "[" & j & "]" & "ĞĞÎ»ºÅ..."
        With xlSheet.Cells(j, rngRef.Column)
            .WrapText = True   '×Ô¶¯»»ĞĞ
            .Rows.AutoFit      '×ÔÊÊÓ¦ĞĞ¸ß
        End With
        '×Ô´Ë´ó¶àÊıÏÔÊ¾Ó¦¸Ã¶¼ÕıÈ·ÁË
        If xlSheet.Cells(j, rngRef.Column).Height > 408 Then
            'ÎŞ·¨ÔÙÔö¼ÓĞĞ¸ßÁË  ×ÔÊÊÓ¦ÁĞ¿í
            xlSheet.Cells(j, rngRef.Column).Columns.AutoFit
        End If
        
        
    'b.Ôª¼şÊıºÍÎ»ºÅÊıÒ»ÖÂ
        Process j * 70 / (usedRow - rngNum.Row) + 20, "¼ì²é" & "[" & j & "]" & "ĞĞÔª¼şÊı..."
        If xlSheet.Cells(j, rngQty.Column) <> "" And xlSheet.Cells(j, rngRef.Column) <> "" Then
           If InStr(xlSheet.Cells(j, rngRef.Column), "ÓÅÏÈ¼¶") = 0 Then
               BomAtom = Split(xlSheet.Cells(j, rngRef.Column), Space(1))
               If CInt(xlSheet.Cells(j, rngQty.Column)) <> (UBound(BomAtom) + 1) Then
                  '²»ÏàµÈÑÕÉ«±ê¼Ç
                  xlSheet.Cells(j, rngQty.Column).Interior.Color = 255 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
                  xlSheet.Cells(j, rngRef.Column).Interior.Color = 255 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
                  xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "Ôª¼şÊıÎ»ºÅÊı²»ÏàµÈ£¡"
                  xlSheet.Cells(j, rngNote.Column).Interior.Color = 255 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
                  xlSheet.Cells(j, rngNote.Column).Font.Size = 10
               End If
           End If
        End If
        
        
    'c.Ìí¼Ó±¸×¢ÏµÁĞ
        Process j * 70 / (usedRow - rngNum.Row) + 20, "¼ì²é" & "[" & j & "]" & "ĞĞ¸¨ÖúËµÃ÷..."
        'c.1.FlashÔö¼Ó±¸×¢: ĞèÔ¤±à³Ì
        If InStr(xlSheet.Cells(j, rngDcp.Column), "FLASH") > 0 And _
           InStr(xlSheet.Cells(j, rngNote.Column), "Ô¤±à³Ì") = 0 Then
            xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "ĞèÔ¤±à³Ì"
            xlSheet.Cells(j, rngNote.Column).Font.Size = 10
            xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
        End If
        'c.2.ÉÕÂ¼Èí¼şÌí¼Ó±¸×¢: SMTÓÃ
        If InStr(xlSheet.Cells(j, rngDcp.Column), "ÉÕÂ¼Èí¼ş") > 0 And _
           InStr(xlSheet.Cells(j, rngNote.Column), "ÓÃ") = 0 Then
            xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "SMTÓÃ"
            xlSheet.Cells(j, rngNote.Column).Font.Size = 10
            xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
        End If
        'c.3.¹¤¾ßÈí¼şÌí¼Ó±¸×¢£º"²âÊÔ½×¶ÎÓÃ"
        If InStr(xlSheet.Cells(j, rngDcp.Column), "¹¤¾ßÈí¼ş") > 0 And _
           InStr(xlSheet.Cells(j, rngNote.Column), "ÓÃ") = 0 Then
            xlSheet.Cells(j, rngNote.Column) = xlSheet.Cells(j, rngNote.Column) + "²âÊÔ½×¶ÎÓÃ"
            xlSheet.Cells(j, rngNote.Column).Font.Size = 10
            xlSheet.Cells(j, rngNote.Column).Font.ColorIndex = 5
        End If
    Next j
        
    Process 95, "ËùÓĞ¼ì²â½áÊø£¡"
    '½áÊøËùÓĞµ÷ÕûºÍ´´½¨ Ìí¼Ó
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾

    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    
    
    Process 100, "¼ì²âÍê³É£¡"
    MsgBox "BOM¸ñÊ½ÕûÀíÍê±Ï£¡" & vbCrLf & vbCrLf & "BOM¼ì²é½áÊø£¡", vbMsgBoxSetForeground + vbOKOnly + vbInformation, "ÌáÊ¾"
    
    Exit Function
    
ErrorHandle:
    
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó

    MsgBox "´ò¿ªExcel¸ñÊ½BOMÎÄ¼şÊ±·¢ÉúÒì³££¡", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
End Function

Function clearRefStr(subRefStr As String) As String
    If subRefStr = "" Then
        Exit Function
    End If
    
    clearRefStrSub subRefStr, vbCrLf
    clearRefStrSub subRefStr, vbCr
    clearRefStrSub subRefStr, vbLf
    
    clearRefStr = subRefStr
    
End Function

Function clearRefStrSub(ByRef tmpRefStr As String, spChar As String)
    
    'vbCrLf -> Space(1)
    Do While InStr(tmpRefStr, spChar) > 0
        tmpRefStr = Replace(tmpRefStr, spChar, Space(1))
    Loop
    'Space(2)->Space(1)
    Do While InStr(tmpRefStr, Space(2))
        tmpRefStr = Replace(tmpRefStr, Space(2), Space(1))
    Loop
    'Î»ÓÚ¿ªÊ¼Î»ÖÃµÄSpace(1)
    If InStr(tmpRefStr, Space(1)) = 1 Then
        tmpRefStr = Replace(tmpRefStr, Space(1), "", 1, 1)
    End If
    'Î»ÓÚ½áÊøÎ»ÖÃµÄSpace(1)
    Do While Right(tmpRefStr, 1) = Space(1)
        tmpRefStr = Left(tmpRefStr, Len(tmpRefStr) - 1)
    Loop
    
End Function
