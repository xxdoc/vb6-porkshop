VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling036_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling036_8"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

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

Private m_AparName As Collection
Private m_StockGroup As Collection
Private m_StockGroupArrange As Collection
Private m_SaleAmounts As Collection
Private m_SaleAmountExs As Collection

Private m_TagetDetails As Collection
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
   
   Set m_SaleAmounts = New Collection
   Set m_SaleAmountExs = New Collection
   
   Set m_AparName = New Collection
   Set m_StockGroup = New Collection
   Set m_StockGroupArrange = New Collection
   Set m_TagetDetails = New Collection
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
   
   Set m_SaleAmounts = Nothing
   Set m_SaleAmountExs = Nothing
   
   Set m_AparName = Nothing
   Set m_StockGroup = Nothing
   Set m_StockGroupArrange = Nothing
   Set m_TagetDetails = Nothing
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
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long
Dim FromDate   As Date
Dim ToDate  As Date
Dim TempData As CBillingDoc

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(1, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "ลำดับที่", ">")
   Call m_Headers0.AddField(Cf)

    If mcolParam("SHORT_NAME") = 0 Then
      Call Cf.SetFieldValue(6, "^", "", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(6, "^", "ลูกค้า", "<")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(4, "^", "", "<")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(4, "^", "ลูกค้า", "<")
      Call m_Headers0.AddField(Cf)
   End If

    For Each TempData In m_StockGroup
      If mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
         Call m_Headers1.AddField(Cf)
      End If
      
      If mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", "ยอดขาย", ">")
         Call m_Headers1.AddField(Cf)
      End If
      
      If mcolParam("SHOW_PRICE") > 0 And mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(4, "^", TempData.STOCK_GROUP_NAME, ">")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_PRICE") > 0 Or mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", TempData.STOCK_GROUP_NAME, ">")
         Call m_Headers0.AddField(Cf)
      End If
   Next TempData

   If mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
   End If
      
   If mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "ยอดขาย", ">")
      Call m_Headers1.AddField(Cf)
   End If
      
   If mcolParam("SHOW_PRICE") > 0 And mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(4, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_PRICE") > 0 Or mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   End If

   Set Cf = Nothing

   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
Dim TempFromDate  As Date
Dim TempToDate As Date
   
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

   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = mcolParam("TO_DATE")
   
   Call LoadTagetDetailStockType(m_TagetDetails, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   Call GetDistinctEmpAparDocTypeFree(m_AparName, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"), , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   Call GetDistinctStockGroupDocTypeFree(m_StockGroupArrange, m_TagetDetails, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"), , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   Call ArrangeByStockTypeCode(m_StockGroupArrange, m_StockGroup)
   
   Call GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM2(m_SaleAmounts, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   Call GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM2(m_SaleAmountExs, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(TempFromDate), "N/A")
   ToDate = "ถึงวันที่ " & EmptyToString(DateToStringExtEx2(TempToDate), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
                                       
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
Dim k As Long
Dim L As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_BillingDoc As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim Total2(25) As Double
Dim Total3(25) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String

Dim SumAmount As Double
Dim SumTotalPrice As Double

Dim TempBd As CBillingDoc
Dim TempStockGroup As CBillingDoc
Dim TempStockItem As CBillingDoc
Dim AmountDate As Long
Dim Ratio As Double

Dim FromDate As Date
Dim ToDate As Date
Dim Tg As CTagetDetail
   
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

   For j = 1 To UBound(Total2)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
      
   k = 2
   SumAmount = 0
   SumTotalPrice = 0
   For Each TempStockItem In m_StockGroup
      Set Tg = GetObject("CTagetDetail", m_TagetDetails, TempStockItem.STOCK_TYPE_CODE)
      If mcolParam("SHOW_AMOUNT") > 0 Then
         k = k + 1
         Total1(k) = Tg.TOTAL_AMOUNT
         SumAmount = SumAmount + Tg.TOTAL_AMOUNT
      End If
         
      If mcolParam("SHOW_PRICE") > 0 Then
         k = k + 1
         Total1(k) = Tg.TOTAL_PRICE
         SumTotalPrice = SumTotalPrice + Tg.TOTAL_PRICE
      End If
   Next TempStockItem
      
   If mcolParam("SHOW_AMOUNT") > 0 Then
      k = k + 1
      Total1(k) = Total1(k) + SumAmount
   End If
   If mcolParam("SHOW_PRICE") > 0 Then
      k = k + 1
      Total1(k) = Total1(k) + SumTotalPrice
   End If
   Call GenerateFooter("", "เป้าการขายเดือน " & EmptyToString(mcolParam("MONTH_ID"), "N/A"), Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)

   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j

      L = 0
      I = 0
      For Each m_BillingDoc In m_AparName
         Call m_Details1.ClearField
         I = I + 1

         If PrevKey1 <> m_BillingDoc.APAR_GROUP_NAME And I <> 1 Then
               If mcolParam("NOT_SHOW_CUSTOMERS") = 0 Then
                  Call GenerateFooter("", "รวม : " & PrevKey1, Total2, m_Details1, 0)
                  Vsp.FontBold = True
               Else
                  L = L + 1
                  Call GenerateFooter("" & L, "กลุ่มลูกค้า : " & PrevKey1, Total2, m_Details1, 0)
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False

               Call m_Details1.ClearField

               For j = 1 To UBound(Total2)
                  Total2(j) = 0
               Next j

               If mcolParam("NOT_SHOW_CUSTOMERS") = 0 Then
                  Vsp = "กลุ่มลูกค้า : " & m_BillingDoc.APAR_GROUP_NAME
                  I = 1
               End If
         ElseIf I = 1 And mcolParam("NOT_SHOW_CUSTOMERS") = 0 Then
               Vsp = "กลุ่มลูกค้า : " & m_BillingDoc.APAR_GROUP_NAME
         End If

         PrevKey1 = m_BillingDoc.APAR_GROUP_NAME

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         If mcolParam("SHORT_NAME") = 0 Then
            TempStr = m_BillingDoc.APAR_NAME & " / " & m_BillingDoc.APAR_CODE
         Else
            TempStr = m_BillingDoc.APAR_SHORT_NAME & " / " & m_BillingDoc.APAR_CODE
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = 2
         SumAmount = 0
         SumTotalPrice = 0
         For Each TempStockItem In m_StockGroup
            If mcolParam("SHOW_AMOUNT") > 0 Then
               k = k + 1
               Amt = 0
               Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & INVOICE_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RECEIPT1_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RETURN_DOCTYPE)
               Amt = Amt - TempBd.TOTAL_AMOUNT
         
               Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & INVOICE_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RECEIPT1_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RETURN_DOCTYPE)
               Amt = Amt - TempBd.TOTAL_AMOUNT
               
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               SumAmount = SumAmount + Amt
   
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            End If
            
            If mcolParam("SHOW_PRICE") > 0 Then
               k = k + 1
               Amt = 0
               Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & INVOICE_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RECEIPT1_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RETURN_DOCTYPE)
               Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
   
               Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & INVOICE_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RECEIPT1_DOCTYPE)
               Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
               Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.APAR_CODE & "-" & TempStockItem.STOCK_TYPE_CODE & "-" & RETURN_DOCTYPE)
               Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
   
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               SumTotalPrice = SumTotalPrice + Amt
   
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
            End If
         Next TempStockItem
         If mcolParam("SHOW_AMOUNT") > 0 Then
            k = k + 1
            Total1(k) = Total1(k) + SumAmount
            Total2(k) = Total2(k) + SumAmount
   
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(SumAmount)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         
         If mcolParam("SHOW_PRICE") > 0 Then
            k = k + 1
            Total1(k) = Total1(k) + SumTotalPrice
            Total2(k) = Total2(k) + SumTotalPrice
   
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(SumTotalPrice)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If

         If mcolParam("NOT_SHOW_CUSTOMERS") = 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      Next m_BillingDoc

      If mcolParam("NOT_SHOW_CUSTOMERS") = 0 Then
            Call GenerateFooter("", "รวม : " & PrevKey1, Total2, m_Details1, 0)
            Vsp.FontBold = True
      Else
            L = L + 1
            Call GenerateFooter("" & L, "กลุ่มลูกค้า : " & PrevKey1, Total2, m_Details1, 0)
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("", "สุทธิ", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateEndDoc
      
      Set m_SaleAmounts = Nothing
      Set m_SaleAmountExs = Nothing
      
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

Private Sub GenerateFooter(txt1 As String, txt2 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j >= 2 Then
         TempStr = FormatNumberToNull(Tot(j))
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
Private Function GetKey1(Tg As CBillingDoc) As String
   GetKey1 = Trim(Tg.APAR_CODE)
End Function
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub
Private Sub ArrangeByStockTypeCode(TempColl As Collection, OutPutColl As Collection)
'On Error Resume Next
Dim StockTypeCode As CBillingDoc
Dim Min As String
Dim TempCode As String
Dim I As Long
   If TempColl.Count <= 0 Then
      Exit Sub
   End If
   
   I = 0
   For Each StockTypeCode In TempColl
      I = I + 1
      If I = 1 Then
         Min = StockTypeCode.STOCK_TYPE_CODE
         TempCode = StockTypeCode.STOCK_TYPE_CODE
      End If
      If StockTypeCode.STOCK_TYPE_CODE <= Min Then
         Min = StockTypeCode.STOCK_TYPE_CODE
         TempCode = StockTypeCode.STOCK_TYPE_CODE
      End If
   Next StockTypeCode
   Set StockTypeCode = Nothing
   
   Call OutPutColl.add(TempColl.Item(TempCode))
   Call TempColl.Remove(TempCode)
   
   Call ArrangeByStockTypeCode(TempColl, OutPutColl)
   
End Sub
