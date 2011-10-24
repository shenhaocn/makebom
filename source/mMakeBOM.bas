Attribute VB_Name = "mMakeBOM"
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
        If insertStr(BMF_TP1) = "-" Then
            xlSheet.Cells(ItemNum + Row, 9) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 9) = insertStr(BMF_TP1)
            If insertStr(BMF_TP1) = "0" Or InStr(insertStr(BMF_TP1), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 9).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
        
        If insertStr(BMF_TP2) = "-" Then
            xlSheet.Cells(ItemNum + Row, 10) = ""
        Else
            xlSheet.Cells(ItemNum + Row, 10) = insertStr(BMF_TP2)
            If insertStr(BMF_TP2) = "0" Or InStr(insertStr(BMF_TP2), "-") = 1 Then
                xlSheet.Cells(ItemNum + Row, 10).Interior.Color = 52479 'ÒÔÇ¿µ÷µÄÑÕÉ«ÏÔÊ¾
            End If
        End If
            
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
        'PCBA_BOM
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
        
        'Ìî³äÊı¾İ
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
    Case BOM_NCDBG:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_NC_DBG.xls")
    Case BOM_NONE:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_None_PartRef.xls")
    Case BOM_Ô¤:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_Ô¤BOM_BMF.xls")
    Case BOM_ÁìÁÏ:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_ÁìÁÏBOM.xls")
    Case BOM_µ÷ÊÔ:
        Set xlBook = xlApp.Workbooks.Open(SaveAsPath & "_µ÷ÊÔBOM.xls")
    Case BOM_Éú²ú:
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
        Case BOM_NCDBG:
            Set rngPos1 = .Find("NCÔª¼ş", lookin:=xlValues)
            Set rngPos2 = .Find("DBGÔª¼ş", lookin:=xlValues)
            Set rngPos3 = .Find("DBG_NCÔª¼ş", lookin:=xlValues)
            
        Case BOM_NONE:
            Set rngPos1 = .Find("None", lookin:=xlValues)
            Set rngPos2 = .Find("None", lookin:=xlValues)
            Set rngPos3 = .Find("None", lookin:=xlValues)
    
        Case BOM_Ô¤, BOM_ÁìÁÏ, BOM_µ÷ÊÔ, BOM_Éú²ú:
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
    Dim bmfBom          As String
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
    If bt_value = BOM_ÁìÁÏ Then
        OrgEnable = True
    Else
        OrgEnable = False
    End If
    
    If Dir(BmfFilePath) = "" Then
        Exit Function
    End If
    
    bmfBom = GetFileContents(BmfFilePath)
    
    bmfBomLine = Split(bmfBom, vbCrLf)

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
            
            If bt_value = BOM_µ÷ÊÔ Or BOM_ÁìÁÏ Or BOM_Éú²ú Or BOM_Ô¤ Then
                '========================================================
                'ÆÕÍ¨Ôª¼ş Çø·ÖÔª¼şÌù×°ÀàĞÍ ÏÈ²»Çø·ÖÕâËÄ¸öBOM ºóĞøµ÷Õû
                Select Case bmfAtom(BMF_MountType)
                Case "S":
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                    
                Case "S+":
                    ItemNum1 = ItemNum1 + 1
                    xlsInsert xlSheet, ItemNum1, rngPos1.Row, bmfAtom, OrgEnable
                    xlSheet.Rows((rngPos1.Row + ItemNum1) & ":" & (rngPos1.Row + ItemNum1)).Interior.Color = 16737792
                    
                Case "L":
                    ItemNum2 = ItemNum2 + 1
                    xlsInsert xlSheet, ItemNum2, rngPos2.Row, bmfAtom, OrgEnable
                    
                Case "N":
                    'Do Nothing
                    
                Case Else
                    '¿âÎÄ¼şÖĞÃ»ÓĞ²éµ½·â×°£¬¾Ü¾øÉú³ÉBOM
                    MsgBox "Î´Öª·â×°[" & bmfAtom(BMF_PcbFB) & "]£¬Çë¸üĞÂ¿âÎÄ¼ş£¡"
                    GoTo ErrorHandle
                End Select
            End If
            
        End If
        
    Next i
    
    If bt_value = BOM_µ÷ÊÔ Or BOM_ÁìÁÏ Or BOM_Éú²ú Or BOM_Ô¤ Then
        'ĞŞ¸Ä»úĞÍÃû³Æ
        xlSheet.Cells(2, 1) = "»úĞÍ£º  " & MainForm.ItemNameText.Text & "            PCBA °æ±¾£º                       °ë³ÉÆ·±àºÅ£º"
        If MainForm.ItemNameText.Text = "" Then
            xlSheet.Cells(2, 1).Font.ColorIndex = 5
        End If
        
    End If
    
    'µ÷ÕûÇø·Ö¸÷ÖÖ²»Í¬µÄBOM
    'Ô¤BOM£ºCaptureÖĞµ¼³öµÄBOM³ıNoneÔª¼ş¡¢NCÔª¼ş¡¢DBGÔª¼ş¡¢DBG_NCÔª¼şÖ®ÍâµÄËùÓĞÔª¼şµÄ¼¯ºÏ
    'ÁìÁÏBOM £º Ô¤BOM + DBGÔª¼ş - ĞÂ´òÑùÎïÁÏ+ ÎïÁÏ¿â´æĞÅÏ¢
    'µ÷ÊÔBOM £º Ô¤BOM + DBGÔª¼ş ¼¯ºÏ
    'Éú²úBOM £º Ô¤BOM + DBG_NCÔª¼ş ¼¯ºÏ
    'Select Case bt_value
    
    'End Select
     
    xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾

    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    
    Exit Function
    
ErrorHandle:

    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    
    MsgBox "Éú³ÉBOMÖĞ¼äÎÄ¼şÊ±·¢ÉúÒì³£", vbCritical + vbMsgBoxSetForeground + vbOKOnly, "´íÎó"
    
End Function


Function BomAdjust() As Boolean
    On Error GoTo ErrorHandle
    
    Dim xlApp As Excel.Application
    Dim ÁìÁÏBOM_xlBook As Excel.Workbook
    Dim ÁìÁÏBOM_xlSheet As Excel.Worksheet
    
    Set xlApp = CreateObject("Excel.Application") '´´½¨EXCEL¶ÔÏó
    xlApp.Visible = False  'ÉèÖÃEXCEL¶ÔÏó¿É¼û£¨»ò²»¿É¼û£©
    
    Set ÁìÁÏBOM_xlBook = xlApp.Workbooks.Open(SaveAsPath & "_ÁìÁÏBOM.xls")
    Set ÁìÁÏBOM_xlSheet = ÁìÁÏBOM_xlBook.Worksheets(1)
    
    'ÔÚÁìÁÏBOMÖĞ²åÈëÁĞ(I£ºTP1¿â´æ) (J£ºTP2¿â´æ) (K£ºTP3¿â´æ)£¨ĞèÑ¡Ôñ¿â´æĞÅÏ¢£©
    'µ÷Õû±í¸ñÊÊÓ¦ĞÂÌí¼ÓµÄÄÚÈİ
    ÁìÁÏBOM_xlSheet.Columns("C:C").ColumnWidth = 45
    ÁìÁÏBOM_xlSheet.Columns("G:G").ColumnWidth = 12
    ÁìÁÏBOM_xlSheet.Columns("H:H").ColumnWidth = 12
    
    ÁìÁÏBOM_xlSheet.Columns("H:H").Copy
    ÁìÁÏBOM_xlSheet.Columns("I:I").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    ÁìÁÏBOM_xlSheet.Columns("I:I").Copy
    ÁìÁÏBOM_xlSheet.Columns("J:J").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    ÁìÁÏBOM_xlSheet.Columns("K:K").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                          SkipBlanks:=False, Transpose:=False
    xlApp.CutCopyMode = False
    
    'Ìî³äÊı¾İ
    ÁìÁÏBOM_xlSheet.Cells(5, 9) = "TP1¿â´æ"
    ÁìÁÏBOM_xlSheet.Cells(5, 10) = "TP2¿â´æ"
    ÁìÁÏBOM_xlSheet.Cells(5, 11) = "TP3¿â´æ"

    ÁìÁÏBOM_xlSheet.Cells(5, 9).Select
    
    Dim LeadPartNum     As Integer
    Dim SmtPartNum      As Integer
    Dim OtherPartNum    As Integer
    
    '»ñÈ¡Ôª¼ş¸öÊı
    LeadPartNum = PartNum(3)
    SmtPartNum = PartNum(4)
    OtherPartNum = PartNum(5)
    
    'É¾³ı´òÑùÔª¼ş
    If DelSamplePart(ÁìÁÏBOM_xlSheet) = False Then
        GoTo ErrorHandle
    End If
    
    ÁìÁÏBOM_xlBook.Save
    
    'ÖØĞÂ±àºÅ
    If ReNum(ÁìÁÏBOM_xlSheet) = False Then
        GoTo ErrorHandle
    End If
        
    ÁìÁÏBOM_xlBook.Save
    ÁìÁÏBOM_xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    
    BomAdjust = True
    Exit Function

ErrorHandle:
    
    ÁìÁÏBOM_xlBook.Close (True) '¹Ø±Õ¹¤×÷²¾
    xlApp.Quit '½áÊøEXCEL¶ÔÏó
    Set xlApp = Nothing 'ÊÍ·ÅxlApp¶ÔÏó
    
    BomAdjust = False
    
End Function

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
    Dim PartNum    As Integer
    Dim DelRows()  As Integer
    
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



