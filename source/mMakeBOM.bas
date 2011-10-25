Attribute VB_Name = "mMakeBOM"
Option Explicit

'BOMÀàÐÍ
Public Enum BomType

BOM_ALL = 0
BOM_NCDBG
BOM_NONE

BOM_Ô¤
BOM_ÁìÁÏ
BOM_µ÷ÊÔ
BOM_Éú²ú

End Enum

'BMFÎÄ¼þ±àÂë¸ñÊ½
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

Function xlsInsert(xlSheet As Excel.Worksheet, ItemNum As Integer, Row As Long, insertStr() As String, OrgEnable As Boolean)
    
    'Ê×ÐÐ²»ÐèÒª¼ÓÈë
    If ItemNum > 1 Then
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Insert
        xlSheet.Rows(Row + ItemNum & ":" & Row + ItemNum).Interior.Pattern = xlNone 'È¥³ýÑÕÉ«µÈ¸ñÊ½ ÐÞÕýÏÔÊ¾bug
    End If
    
    xlSheet.Cells(ItemNum + Row, 1) = ItemNum
    xlSheet.Cells(ItemNum + Row, 2) = insertStr(BMF_PartNum)
    xlSheet.Cells(ItemNum + Row, 3) = insertStr(BMF_Description)
    xlSheet.Cells(ItemNum + Row, 5) = insertStr(BMF_Quantity)
    xlSheet.Cells(ItemNum + Row, 6) = insertStr(BMF_PartRef)
    xlSheet.Cells(ItemNum + Row, 7) = insertStr(BMF_PcbFB)
    xlSheet.Cells(ItemNum + Row, 8) = insertStr(BMF_Value)
    
    'ÊÇ·ñÌí¼Ó¿â´æÐÅÏ¢£¿
    If OrgEnable = True Then
    
        'Ìí¼ÓTP1¿â´æÐÅÏ¢
        If insertStr(BMF_TP1) = "-" Then
            xlSheet.Cells(ItemNum + Row, 9) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 9) = insertStr(BMF_TP1)
            If insertStr(BMF_TP1) = "0" Or InStr(insertStr(BMF_TP1), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 9).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
        
        'Ìí¼ÓTP2¿â´æÐÅÏ¢
        If insertStr(BMF_TP2) = "-" Then
            xlSheet.Cells(ItemNum + Row, 10) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 10) = insertStr(BMF_TP2)
            If insertStr(BMF_TP2) = "0" Or InStr(insertStr(BMF_TP2), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 10).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
            
        'Ìí¼ÓTP3¿â´æÐÅÏ¢
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

'ÅÅÐòÊý×é£¨¿ÉÒÔÊÇ×Ö·û´®Êý×é£©£ºBigToSmall=True ´Ó´óµ½Ð¡£¬·ñÔò ´ÓÐ¡´óµ½
Function ZuSorted(Zu() As Variant, prefixStr As String, Optional BigToSmall As Boolean) As String
   
   Dim i As Long, j As Long, S As Variant
   Dim TF As Boolean, nL As Long, nU As Long
   
   nL = LBound(Zu): nU = UBound(Zu)
   For i = nL To nU
      For j = nL To nU '       ´Ó´óµ½Ð¡                ´ÓÐ¡´óµ½
        If BigToSmall Then TF = Zu(j) < Zu(i) Else TF = Zu(j) > Zu(i)
        If TF Then S = Zu(i): Zu(i) = Zu(j): Zu(j) = S
      Next j
   Next i
   
   For i = nL To nU - 1
      ZuSorted = ZuSorted + prefixStr + CStr(Zu(i)) + Space(1)
   Next i
   
   ZuSorted = ZuSorted + prefixStr + CStr(Zu(nU))
   
End Function

'ÓÉÓÚÎ»ºÅ³¤¶È²»Ò»£¬³£¹æÅÅÐò·½·¨²»¿ÉÐÐ Òò´ËÐèÒªÌØÊâÅÅÐò·½·¨
Function RealSorted(RefStr As String, Optional BigToSmall As Boolean) As String
    Dim srcStr()    As String
    Dim intSorted() As Variant
    Dim i           As Long
    Dim Index       As Long 'Î»ºÅºóÊý×Ö¿ªÊ¼µÄÎ»ÖÃ
    
    Dim prefixStr   As String
    
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
        intSorted(i) = Val(Right(srcStr(i), Len(srcStr(i)) - Index))
    Next i
    
    RealSorted = ZuSorted(intSorted, prefixStr, BigToSmall)
    
End Function

'Ìí¼ÓDBG NCÔª¼þµ½ÏàÓ¦µÄExcel BOMÖÐ ¶ÔºÏ²¢µÄÎ»ºÅ½øÐÐÅÅÐò
Function addDbgNcPart(xlSheet As Excel.Worksheet, bmfAtom() As String, _
                      ByRef ItemNum1 As Integer, ByRef ItemNum2 As Integer, _
                      rngPos1 As Range, rngPos2 As Range, OrgEnable As Boolean)

    Dim rngNum       As Range
    Dim partRefStr() As String
    Dim i            As Integer
    
    'ÊÇ·ñÐèÒªºÏ²¢Ôª¼þ
    With xlSheet.Cells
        Set rngNum = .Find(bmfAtom(BMF_PartNum), lookin:=xlValues)
        If rngNum Is Nothing Then
            '²»ÐèÒªºÏ²¢ Ö±½ÓÌí¼ÓÔÚºóÃæ
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
                '¿âÎÄ¼þÖÐÃ»ÓÐ²éµ½·â×°£¬¾Ü¾øÉú³ÉBOM
                MsgBox "Î´Öª·â×°[" & bmfAtom(BMF_PcbFB) & "]£¬Çë¸üÐÂ¿âÎÄ¼þ£¡"
                
            End Select
            
        Else 'ÐèÒªºÏ²¢ ÐÞ¸ÄÊýÁ¿(µÚ5ÁÐ) ºÍÎ»ºÅ(µÚ6ÁÐ) Î»ºÅÐèÒªÖØÐÂÅÅÐò
            'xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            xlSheet.Cells(rngNum.Row, 5) = CInt(xlSheet.Cells(rngNum.Row, 5)) + CInt(bmfAtom(BMF_Quantity))
            xlSheet.Cells(rngNum.Row, 5).Font.ColorIndex = 5
            'Î»ºÅÐèÒªÅÅÐò£¡
            xlSheet.Cells(rngNum.Row, 6) = RealSorted(xlSheet.Cells(rngNum.Row, 6) + " " + bmfAtom(BMF_PartRef), False)
            xlSheet.Cells(rngNum.Row, 6).Font.ColorIndex = 5
        End If
    End With
End Function

Function CopyLine(xlSheetTo As Excel.Worksheet, RowTo As Integer, xlSheetFrom As Excel.Worksheet, RowFrom As Integer, ColumnNum As Integer, PartNum As Integer)
    xlSheetTo.Rows(RowTo & ":" & RowTo).Insert
    xlSheetTo.Cells(RowTo, 1) = PartNum
    Dim i As Integer
    For i = 2 To ColumnNum
        xlSheetTo.Cells(RowTo, i) = xlSheetFrom.Cells(RowFrom, i)
    Next i
    xlSheetTo.Rows(RowTo & ":" & RowTo).Font.ColorIndex = 5
End Function

'¸ù¾ÝÑ¡Ïî´´½¨BOM ²¢µ÷Õû¸ñÊ½ ÎªÌî³äÊý¾Ý×ö×¼±¸
Function ExcelCreate(bt_value As BomType)

    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    '=====================================================================================
    'Ô¤BOM£ºCaptureÖÐµ¼³öµÄBOM³ýNoneÔª¼þ¡¢NCÔª¼þ¡¢DBGÔª¼þ¡¢DBG_NCÔª¼þÖ®ÍâµÄËùÓÐÔª¼þµÄ¼¯ºÏ¡£
    '=====================================================================================
    If bt_value = BOM_Ô¤ Then
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_Ô¤BOM_BMF.xls")
        
        xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    End If
    
    '=====================================================================================
    'NC_DBGÔª¼þxls
    '=====================================================================================
    If bt_value = BOM_NCDBG Then
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_NC_DBG.xls")
        
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'NoneÔª¼þxls
    '=====================================================================================
    If bt_value = BOM_NONE Then
    
        Dim rngNC           As Range
        Dim rngDB           As Range
        Dim rngDBNC         As Range
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\NC_DBG_template.xls")
        xlBook.SaveAs (SaveAsPath & "_None_PartRef.xls")
        xlBook.Worksheets(1).Name = "NoneÔª¼þ"
        
        Set xlSheet = xlBook.Worksheets(1)
        
        With xlSheet.Cells
            Set rngNC = .Find("NCÔª¼þ", lookin:=xlValues)
            Set rngDB = .Find("DBGÔª¼þ", lookin:=xlValues)
            Set rngDBNC = .Find("DBG_NCÔª¼þ", lookin:=xlValues)
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
    'ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼þ - ÐÂ´òÑùÎïÁÏ+ ÎïÁÏ¿â´æÐÅÏ¢¡£
    '=====================================================================================
    If bt_value = BOM_ÁìÁÏ Then
        
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\template\PCBA_BOM_template.xls")
        Set xlSheet = xlBook.Worksheets(1)
        xlBook.SaveAs (SaveAsPath & "_ÁìÁÏBOM.xls")
         
        'µ÷ÕûÁìÁÏBOM¸ñÊ½
        'ÔÚÁìÁÏBOMÖÐ²åÈëÁÐ(I£ºTP1¿â´æ) (J£ºTP2¿â´æ) (K£ºTP3¿â´æ)£¨ÐèÑ¡Ôñ¿â´æÐÅÏ¢£©
        'µ÷Õû±í¸ñÊÊÓ¦ÐÂÌí¼ÓµÄÄÚÈÝ
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
        
        'Ìî³äÊý¾Ý µÚ5ÐÐ²ÅÊÇÕæÕýµÄÁÐÍ·ÃèÊöÐÐ
        xlSheet.Cells(5, 9) = "TP1¿â´æ"
        xlSheet.Cells(5, 10) = "TP2¿â´æ"
        xlSheet.Cells(5, 11) = "TP3¿â´æ"

        xlSheet.Cells(5, 1).Select
         
        xlBook.Close (True)
    End If
    
    '=====================================================================================
    'µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼þ ¼¯ºÏ¡£
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
    'Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼þ ¼¯ºÏ
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
    
    MsgBox "´´½¨BOMÖÐ¼äÎÄ¼þÊ±·¢ÉúÒì³£", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
    

End Function

'Ìî³äÏàÓ¦µÄÊý¾Ý

'=====================================================================================
'1.Ô¤BOM£ºCaptureÖÐµ¼³öµÄBOM³ýNoneÔª¼þ¡¢NCÔª¼þ¡¢DBGÔª¼þ¡¢DBG_NCÔª¼þÖ®ÍâµÄËùÓÐÔª¼þµÄ¼¯ºÏ¡£
'2.NC_DBGÔª¼þxls
'3.NoneÔª¼þxls
'4.ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼þ - ÐÂ´òÑùÎïÁÏ+ ÎïÁÏ¿â´æÐÅÏ¢
'5.µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼þ ¼¯ºÏ
'6.Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼þ ¼¯ºÏ
'ÿ
'Ôª¼þÀàÐÍ £º ÆÕÍ¨Ôª¼þ NcDbgÔª¼þ NoneÔª¼þ ´òÑùÎïÁÏ ÿ
'=====================================================================================
Function CreateBOM(bt_value As BomType) As Boolean

    On Error GoTo ErrorHandle

    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False

    '´ò¿ªÏàÓ¦µÄÎÄ¼þ
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

    '¶¨Î»¸÷ÖÖÔª¼þÎ»ÖÃ
    Dim rngPos1       As Range 'ÔÚBOMÎÄ¼þÖÐ±íÊ¾"SMTÔª¼þ"Î»ÖÃ ÔÚDBG¡¢NoneÔª¼þ´ú±í"NCÔª¼þ"Î»ÖÃ
    Dim rngPos2       As Range 'ÔÚBOMÎÄ¼þÖÐ±íÊ¾"DIPÔª¼þ"Î»ÖÃ ÔÚDBG¡¢NoneÔª¼þ´ú±í"DBGÔª¼þ"Î»ÖÃ
    Dim rngPos3       As Range 'ÔÚBOMÎÄ¼þÖÐ±íÊ¾"ÆäËûÔª¼þ"Î»ÖÃ ÔÚDBG¡¢NoneÔª¼þ´ú±í"DBG_NCÔª¼þ"Î»ÖÃ
    With xlSheet.Cells

        Select Case bt_value
            Case BOM_NCDBG

            Set rngPos1 = .Find("NCÔª¼þ", lookin:=xlValues)
            Set rngPos2 = .Find("DBGÔª¼þ", lookin:=xlValues)
            Set rngPos3 = .Find("DBG_NCÔª¼þ", lookin:=xlValues)

            Case BOM_NONE

            Set rngPos1 = .Find("None", lookin:=xlValues)
            Set rngPos2 = .Find("None", lookin:=xlValues)
            Set rngPos3 = .Find("None", lookin:=xlValues)

            Case BOM_Ô¤, BOM_ÁìÁÏ, BOM_µ÷ÊÔ, BOM_Éú²ú

            Set rngPos1 = .Find("SMTÔª¼þ", lookin:=xlValues)
            Set rngPos2 = .Find("DIPÔª¼þ", lookin:=xlValues)
            Set rngPos3 = .Find("ÆäËûÔª¼þ", lookin:=xlValues)

            Case Else
            GoTo ErrorHandle
        End Select

        If rngPos1 Is Nothing Or rngPos2 Is Nothing Or rngPos3 Is Nothing Then
            MsgBox "Ä£°å´íÎó-Ôª¼þÎ»ÖÃ¶¨Î»´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
            GoTo ErrorHandle
        End If
    End With

    '=======================================================
    '¿ªÊ¼¶ÁÈ¡bmfÎÄ¼þ²¢Ìî³äÏàÓ¦µÄÄÚÈÝ
    '=======================================================
    Dim bmfBomLine()    As String
    Dim bmfAtom()       As String

    Dim ItemNum1        As Integer 'NCÔª¼þ »ò NoneÔª¼þ »ò SMTÔª¼þÊý
    Dim ItemNum2        As Integer 'DBGÔª¼þ »ò NoneÔª¼þ »ò DIPÔª¼þÊý
    Dim ItemNum3        As Integer 'DBG_NCÔª¼þ »ò NoneÔª¼þ »ò ÆäËûÔª¼þÊý

    Dim i               As Integer
    Dim OrgEnable       As Boolean

    ItemNum1 = 0
    ItemNum2 = 0
    ItemNum3 = 0

    'ÊÇ·ñÌí¼Ó¿â´æÐÅÏ¢£¿
    'ÐèÒª¼ì²âtsvÎÄ¼þ´´½¨Ê±¼ä Èç¹ûÊ±¼ä²»ÔÚ3ÌìÄÚ ¿â´æÐÅÏ¢Ìí¼ÓÁËÒ²ÊÇÃ»ÓÐÓÃµÄ
    'ÒÑ¾­±¾Functionµ÷ÓÃÇ°¼ì²âÁËtsvÎÄ¼þÊ±¼ä
    If bt_value = BOM_ÁìÁÏ Then
        OrgEnable = True
    Else
        OrgEnable = False
    End If

    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If

    '¶ÁÈ¡bmfÎÄ¼þ ²¢°´ÐÐ·Ö¸îÎªÊý×é
    bmfBomLine = Split(GetFileContents(BmfFilePath), vbCrLf)

    '±éÀúbmfÎÄ¼þ
    For i = 1 To UBound(bmfBomLine) - 1
        bmfAtom = Split(bmfBomLine(i), vbTab)

        If bmfAtom(BMF_PcbFB) = "" Or bmfAtom(BMF_Value) = "" Then
            MsgBox "µÚ" & bmfAtom(BMF_ItemNum) & "ÏîÔª¼þÐÅÏ¢²»ÍêÕû", vbExclamation + vbMsgBoxSetForeground + vbOKOnly, "¾¯¸æ"
            GoTo ErrorHandle
        End If

        If InStr(bmfAtom(BMF_Value), "_DBG_NC") > 0 Or bmfAtom(BMF_Value) = "DBG_NC" Then
            If bt_value = BOM_NCDBG Then
                'DBG_NCÔª¼þ
                ItemNum3 = ItemNum3 + 1
                xlsInsert xlSheet, ItemNum3, rngPos3.Row, bmfAtom, OrgEnable
            End If

        ElseIf InStr(bmfAtom(BMF_Value), "_DBG") > 0 Or bmfAtom(BMF_Value) = "DBG" Then
            If bt_value = BOM_NCDBG Then
                'DBGÔª¼þ
                ItemNum2 = ItemNum2 + 1
                xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
            End If

        ElseIf InStr(bmfAtom(BMF_Value), "_NC") > 0 Or bmfAtom(BMF_Value) = "NC" Then
            If bt_value = BOM_NCDBG Then
                'NCÔª¼þ
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
                'ÆÕÍ¨Ôª¼þ Çø·ÖÔª¼þÌù×°ÀàÐÍ ÏÈ²»Çø·ÖÕâËÄ¸öBOM ºóÐøµ÷Õû
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
                    '¿âÎÄ¼þÖÐÃ»ÓÐ²éµ½·â×°£¬¾Ü¾øÉú³ÉBOM
                    MsgBox "Î´Öª·â×°[" & bmfAtom(BMF_PcbFB) & "]£¬Çë¸üÐÂ¿âÎÄ¼þ£¡"
                    GoTo ErrorHandle
                    
                End Select
            End If

        End If

    Next i

    If bt_value = BOM_µ÷ÊÔ Or bt_value = BOM_ÁìÁÏ Or _
       bt_value = BOM_Éú²ú Or bt_value = BOM_Ô¤ Then
        'ÐÞ¸Ä»úÐÍÃû³Æ
        xlSheet.Cells(2, 1) = "»úÐÍ£º  " & MainForm.ItemNameText.Text & "            PCBA °æ±¾£º                       °ë³ÉÆ·±àºÅ£º"
        If MainForm.ItemNameText.Text = "" Then
            xlSheet.Cells(2, 1).Font.ColorIndex = 5
        End If

    End If

    If bt_value = BOM_µ÷ÊÔ Or bt_value = BOM_ÁìÁÏ Or bt_value = BOM_Éú²ú Then
        'µ÷ÕûÇø·Ö¸÷ÖÖ²»Í¬µÄBOM
        'Ô¤BOM£ºCaptureÖÐµ¼³öµÄBOM³ýNoneÔª¼þ¡¢NCÔª¼þ¡¢DBGÔª¼þ¡¢DBG_NCÔª¼þÖ®ÍâµÄËùÓÐÔª¼þµÄ¼¯ºÏ
        'ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼þ - ÐÂ´òÑùÎïÁÏ + ÎïÁÏ¿â´æÐÅÏ¢
        'µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼þ ¼¯ºÏ
        'Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼þ ¼¯ºÏ
        
        'ÖØÐÂ±éÀúbmfÎÄ¼þ ¸ù¾Ýµ÷ÕûÏàÓ¦BOMÐÅÏ¢
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
        
        'É¾³ý´òÑùÔª¼þ
        If bt_value = BOM_ÁìÁÏ Then
            If DelSamplePart(xlSheet) = False Then
                GoTo ErrorHandle
            End If
        End If
    
        'ÖØÐÂ±àºÅ
        If ReNum(xlSheet) = False Then
            GoTo ErrorHandle
        End If
    End If
    
    '¼ÆËã»ñÈ¡ExcelµÄ"Î»ºÅÇøÓò"ÐÐ¸ß£¬¿´ÊÇ·ñ·ûºÏÒªÇó£¬²»·ûºÏÐèÒªµ÷ÕûÁÐ¿í
    
    
    '½áÊøËùÓÐµ÷ÕûºÍ´´½¨ Ìí¼Ó
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾

    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó

    Exit Function

ErrorHandle:
    
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó

    MsgBox "Éú³ÉBOMÖÐ¼äÎÄ¼þÊ±·¢ÉúÒì³£", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"

End Function

Function DelSamplePart(xlSheet As Excel.Worksheet) As Boolean
    Dim rngStart        As Range
    Dim rngEND          As Range
    
    'É¾³ý´òìÈÎïÁÏ ÁÏºÅ´æÔÚ µ«ÊÇÊôÓÚ12345xxxxx »òxxxxx xxxxxÀàÐÍ
    With xlSheet.Cells
        Set rngStart = .Find("SMTÔª¼þ", lookin:=xlValues)
        Set rngEND = .Find("END", lookin:=xlValues)
        If rngStart Is Nothing Or rngEND Is Nothing Then
            MsgBox "PCBA_BOMÄ£°å´íÎó", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
            DelSamplePart = False
        End If
    End With
    
    Dim i          As Integer
    Dim j          As Integer
    Dim PartNum    As Integer   'Î±Ôª¼þ¸öÊý
    Dim DelRows()  As Integer   '¼ÇÂ¼ÒªÉ¾³ýµÄÐÐºÅ
    
    PartNum = rngEND.Row - rngStart.Row
    ReDim DelRows(PartNum) As Integer
    
    j = 0
    For i = rngStart.Row To rngEND.Row
        If IsNumeric(xlSheet.Cells(i, 2)) = False _
             And xlSheet.Cells(i, 2) <> "SMTÔª¼þ" _
             And xlSheet.Cells(i, 2) <> "DIPÔª¼þ" _
             And xlSheet.Cells(i, 2) <> "ÆäËûÔª¼þ" _
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

Function ReNum(xlSheet As Excel.Worksheet) As Boolean
    'ÖØÐÂ±àºÅ
    Dim rngSMT          As Range
    Dim rngLEAD         As Range
    Dim rngOther        As Range
    Dim rngEND          As Range
    
    'É¾³ý´òìÈÎïÁÏ ÁÏºÅ´æÔÚ µ«ÊÇÊôÓÚ12345xxxxx »òxxxxx xxxxxÀàÐÍ
    With xlSheet.Cells
        Set rngSMT = .Find("SMTÔª¼þ", lookin:=xlValues)
        Set rngLEAD = .Find("DIPÔª¼þ", lookin:=xlValues)
        Set rngOther = .Find("ÆäËûÔª¼þ", lookin:=xlValues)
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
