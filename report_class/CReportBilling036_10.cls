VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling036_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling036_10"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Dim TotalMoves As Collection
Dim TotalSalePrice As Double



' note: this API is declared incorrectly in the VB API Viewer.


Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
   
   Set TotalMoves = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set TotalMoves = New Collection
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim Amt As Double
Dim iCount As Long

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "ลำดับ", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รหัสสินค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "รายละเอียด", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "หน่วยนับ", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(3, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
   
   ' note: SetTextCharacterExtra takes a spacing in pixels, which
    '       we get by converting from twips.
    'g_CharSpacing = 0 / Vsp.TwipsPerPixelX
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Vsp.ZoomMode = zmPageWidth
   
   Dim TempColl As Collection
   Dim TotalStockCodeCost As Collection
   Dim Count As Long
   Set TempColl = New Collection
   Set TotalStockCodeCost = New Collection
   
   If mcolParam("ORDER_BY") <= 0 Then
      Call GetSaleAmountStockCode2(TotalMoves, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), , TotalSalePrice, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("INCLUDE_FREE"), mcolParam("SALE_CODE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"))
   Else
      Call GetSaleAmountStockCode2(TempColl, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), , TotalSalePrice, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("INCLUDE_FREE"), mcolParam("SALE_CODE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"))
      If mcolParam("ORDER_BY") = 1 And mcolParam("ORDER_TYPE") = 1 Then
         Call ArrangeByAmountMin(TempColl, TotalMoves)
      ElseIf mcolParam("ORDER_BY") = 1 Then
         Call ArrangeByAmountMax(TempColl, TotalMoves)
      ElseIf mcolParam("ORDER_BY") = 2 And mcolParam("ORDER_TYPE") = 1 Then
         Call ArrangeBySaleMin(TempColl, TotalMoves)
      ElseIf mcolParam("ORDER_BY") = 2 Then
         Call ArrangeBySaleMax(TempColl, TotalMoves)
      End If
   End If
   Set TempColl = Nothing
   Set TotalStockCodeCost = Nothing
   
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " พนักงานขาย " & EmptyToString(mcolParam("SALE_CODE"), "N/A") & " เรียงตาม " & EmptyToString(mcolParam("ORDER_BY_NAME"), "N/A") & " เรียงจาก " & EmptyToString(mcolParam("ORDER_TYPE_NAME"), "N/A")
                                       
                                       
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim I As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(20) As Double
Dim iCount As Long
Dim TempStr As String
Dim TotalMove As CCreditStockCode
   
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField
      
         
      I = 0

      For Each TotalMove In TotalMoves
         I = I + 1
         
         Amt = 0
         
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = TotalMove.STOCK_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = TotalMove.STOCK_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = TotalMove.UNIT_CHANGE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(5)
         Amt = TotalMove.TOTAL_AMOUNT
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + Amt
         
         Amt = TotalMove.CREDIT_BALANCE
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + Amt
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
      Next TotalMove
      
      Call GenerateFooter("สุทธิ", "", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateEndDoc
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function
Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempProfit As Double
Dim TempRaw As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j = 10 Then
         TempStr = FormatNumberToNull(MyDiff(TempProfit, TempRaw) * 100)
      ElseIf j >= 5 Then
         TempStr = FormatNumberToNull(Tot(j))
      End If
      
      If j = 6 Then
         TempProfit = Tot(j)
      ElseIf j = 9 Then
         TempRaw = Tot(j)
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long
   
   
   
   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.TAX_ID)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Private Sub ArrangeByAmountMax(TempColl As Collection, OutPutColl As Collection)
On Error GoTo err
Dim CreditStockCode As CCreditStockCode
Dim Max As Double
Dim TempCode As String
Dim I As Long
   Max = 0
   
   If TempColl.Count <= 0 Then
      Exit Sub
   End If
   
   I = 0
   For Each CreditStockCode In TempColl
      I = I + 1
      If I = 1 Then
         Max = CreditStockCode.TOTAL_AMOUNT
         TempCode = CreditStockCode.STOCK_NO
      End If
      If CreditStockCode.TOTAL_AMOUNT >= Max Then
         Max = CreditStockCode.TOTAL_AMOUNT
         TempCode = CreditStockCode.STOCK_NO
      End If
   Next
   Set CreditStockCode = Nothing
   
   Call OutPutColl.add(TempColl.Item(TempCode))
   Call TempColl.Remove(TempCode)
   
   Call ArrangeByAmountMax(TempColl, OutPutColl)
   
   Exit Sub
err:
   'debug.print
End Sub
Private Sub ArrangeBySaleMax(TempColl As Collection, OutPutColl As Collection)
'On Error Resume Next
Dim CreditStockCode As CCreditStockCode
Dim Max As Double
Dim TempCode As String
Dim I As Long
   Max = 0
   
   If TempColl.Count <= 0 Then
      Exit Sub
   End If
   I = 0
   For Each CreditStockCode In TempColl
      I = I + 1
      If I = 1 Then
         Max = CreditStockCode.CREDIT_BALANCE
         TempCode = CreditStockCode.STOCK_NO
      End If
      If CreditStockCode.CREDIT_BALANCE >= Max Then
         Max = CreditStockCode.CREDIT_BALANCE
         TempCode = CreditStockCode.STOCK_NO
      End If
   Next
   Set CreditStockCode = Nothing
   
   Call OutPutColl.add(TempColl.Item(TempCode))
   Call TempColl.Remove(TempCode)
   
   Call ArrangeBySaleMax(TempColl, OutPutColl)
   
End Sub
Private Sub ArrangeByAmountMin(TempColl As Collection, OutPutColl As Collection)
'On Error Resume Next
Dim CreditStockCode As CCreditStockCode
Dim Min As Double
Dim TempCode As String
Dim I As Long
   If TempColl.Count <= 0 Then
      Exit Sub
   End If
   
   I = 0
   For Each CreditStockCode In TempColl
      I = I + 1
      If I = 1 Then
         Min = CreditStockCode.TOTAL_AMOUNT
         TempCode = CreditStockCode.STOCK_NO
      End If
      If CreditStockCode.TOTAL_AMOUNT <= Min Then
         Min = CreditStockCode.TOTAL_AMOUNT
         TempCode = CreditStockCode.STOCK_NO
      End If
   Next CreditStockCode
   Set CreditStockCode = Nothing
   
   Call OutPutColl.add(TempColl.Item(TempCode))
   Call TempColl.Remove(TempCode)
   
   Call ArrangeByAmountMin(TempColl, OutPutColl)
   
End Sub
Private Sub ArrangeBySaleMin(TempColl As Collection, OutPutColl As Collection)
'On Error Resume Next
Dim CreditStockCode As CCreditStockCode
Dim Min As Double
Dim TempCode As String
Dim I As Long
   If TempColl.Count <= 0 Then
      Exit Sub
   End If
   
   I = 0
   For Each CreditStockCode In TempColl
      I = I + 1
      If I = 1 Then
         Min = CreditStockCode.CREDIT_BALANCE
         TempCode = CreditStockCode.STOCK_NO
      End If
      If CreditStockCode.CREDIT_BALANCE <= Min Then
         Min = CreditStockCode.CREDIT_BALANCE
         TempCode = CreditStockCode.STOCK_NO
      End If
   Next
   Set CreditStockCode = Nothing
   
   Call OutPutColl.add(TempColl.Item(TempCode))
   Call TempColl.Remove(TempCode)
   
   Call ArrangeBySaleMin(TempColl, OutPutColl)
   
End Sub
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub
