VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportTaget002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportTaget002"
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

Private m_DistinctMain As Collection
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
   
   Set m_DistinctMain = New Collection
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
   
   Set m_DistinctMain = Nothing
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
   
   Call Cf.SetFieldValue(6, "^", "รายละเอียด", "<")
   Call m_Headers1.AddField(Cf)
   
   '------------------------------------------------------------------------------------------>          ยอดขาย
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   
   '------------------------------------------------------------------------------------------>
   
   '------------------------------------------------------------------------------------------>          รับคืน
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ปริมาณ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   
   '------------------------------------------------------------------------------------------>
   
   Call Cf.SetFieldValue(6, "<", "", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "TAGET", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "ACTUAL", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "DIFF", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "TAGET", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "ACTUAL", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "DIFF", "<")
   Call m_Headers0.AddField(Cf)
   
   
   Call Cf.SetFieldValue(6, "^", "", "<")
   Call m_Headers1_1.AddField(Cf)
   
   Call Cf.SetFieldValue(16, "^", "ยอดขาย", ">")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(16, "^", "รับคืน", ">")
   Call m_Headers1_1.AddField(Cf)
   
   
   Set Cf = Nothing
   
   Call m_Headers1_1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
   
   ' note: SetTextCharacterExtra takes a spacing in pixels, which
    '       we get by converting from twips.
    'g_CharSpacing = -0 / Vsp.TwipsPerPixelX
   
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
   
   Call GetDistinctBranchEmpStockCodeDocTypeFree(m_DistinctMain, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   
   Call LoadTagetDetailBranchEmpStock(m_TagetDetails, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   
   Call GetSaleAmountBranchEmpStockCodeDocTypeFree(m_SaleAmounts, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   Call GetSaleAmountBranchEmpStockCodeDocTypeFreeEx(m_SaleAmountExs, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate
                                       
                                       
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
Dim m_BillingDoc As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim Total2(25) As Double
Dim Total3(25) As Double
Dim Total4(25) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String

Dim TempAmount1 As Double
Dim TempTotalPrice1 As Double
Dim TempAmount2 As Double
Dim TempTotalPrice2 As Double
Dim TempBd As CBillingDoc

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
      Total4(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField
      
      AmountDate = DateDiff("D", mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE")) + 1
      
      Ratio = MyDiffEx(AmountDate, GetFirstLastDateEX(mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), FromDate, ToDate))
      
      I = 0
      
      For Each m_BillingDoc In m_DistinctMain
         I = I + 1
         
         If PrevKey1 <> m_BillingDoc.CUSTOMER_BRANCH_CODE And I > 1 Then
            
            Call GenerateFooter("      รวม" & PrevKey3, Total2, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            Call GenerateFooter("รวม เขตการขาย   : " & PrevKey2 & "", Total3, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
               Total3(j) = 0
            Next j
            
            Vsp = ""
            
            Vsp = "เขตการขาย        : " & m_BillingDoc.CUSTOMER_BRANCH_NAME & " / " & m_BillingDoc.CUSTOMER_BRANCH_CODE
            Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            
         ElseIf PrevKey3 <> m_BillingDoc.SALE_CODE And I <> 1 Then
            
            Call GenerateFooter("      รวม" & PrevKey3, Total2, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
            Next j
            
            Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            
         ElseIf I = 1 Then
            Vsp = "เขตการขาย        : " & m_BillingDoc.CUSTOMER_BRANCH_NAME & " / " & m_BillingDoc.CUSTOMER_BRANCH_CODE
            Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
         End If
         
         PrevKey1 = m_BillingDoc.CUSTOMER_BRANCH_CODE
         PrevKey2 = m_BillingDoc.CUSTOMER_BRANCH_NAME
         PrevKey3 = m_BillingDoc.SALE_CODE
         PrevKey4 = m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
         
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_BillingDoc.STOCK_DESC & "/" & m_BillingDoc.STOCK_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set Tg = GetObject("CTagetDetail", m_TagetDetails, GetKey1(m_BillingDoc))
         Amt = Tg.TOTAL_AMOUNT * Ratio
         If Tg.Flag = "Y" Then
            Amt = 0
         End If
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(2) = Total1(2) + Amt
         Total2(2) = Total2(2) + Amt
         Total3(2) = Total3(2) + Amt
         Total4(2) = Total4(2) + Amt
         TempAmount1 = Amt
         
         Amt = Tg.TOTAL_PRICE * Ratio
         If Tg.Flag = "Y" Then
            Amt = 0
         End If
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(3) = Total1(3) + Amt
         Total2(3) = Total2(3) + Amt
         Total3(3) = Total3(3) + Amt
         Total4(3) = Total4(3) + Amt
         TempTotalPrice1 = Amt
         
         If Amt <> 0 Then
            Tg.Flag = "Y"
         End If
         
         Amt = 0
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_AMOUNT
         
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + Amt
         Total2(4) = Total2(4) + Amt
         Total3(4) = Total3(4) + Amt
         Total4(4) = Total4(4) + Amt
         TempAmount2 = Amt
         
         Amt = 0
         Amt = MyDiff(TempAmount2 * 100, TempAmount1)
         
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(5) = Total1(5) + Amt
         Total2(5) = Total2(5) + Amt
         Total3(5) = Total3(5) + Amt
         Total4(5) = Total4(5) + Amt
         
         Amt = 0
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + Amt
         Total2(6) = Total2(6) + Amt
         Total3(6) = Total3(6) + Amt
         Total4(6) = Total4(6) + Amt
         TempTotalPrice2 = Amt
         
         Amt = 0
         Amt = MyDiff(TempTotalPrice2 * 100, TempTotalPrice1)
         
          Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + Amt
         Total2(7) = Total2(7) + Amt
         Total3(7) = Total3(7) + Amt
         Total4(7) = Total4(7) + Amt
          
         Amt = TempAmount2 - TempAmount1
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(8) = Total1(8) + Amt
         Total2(8) = Total2(8) + Amt
         Total3(8) = Total3(8) + Amt
         Total4(8) = Total4(8) + Amt
         
         Amt = (MyDiff(TempAmount2, TempAmount1) - 1) * 100
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + Amt
         Total2(9) = Total2(9) + Amt
         Total3(9) = Total3(9) + Amt
         Total4(9) = Total4(9) + Amt
         
         Amt = TempTotalPrice2 - TempTotalPrice1
         
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(10) = Total1(10) + Amt
         Total2(10) = Total2(10) + Amt
         Total3(10) = Total3(10) + Amt
         Total4(10) = Total4(10) + Amt
         
         Amt = (MyDiff(TempTotalPrice2, TempTotalPrice1) - 1) * 100
         
         Set HeadCf = m_Headers1.Fields(11)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(11) = Total1(11) + Amt
         Total2(11) = Total2(11) + Amt
         Total3(11) = Total3(11) + Amt
         Total4(11) = Total4(11) + Amt
         
         Amt = Tg.TOTAL_AMOUNT_RT * Ratio
         Set HeadCf = m_Headers1.Fields(12)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(12) = Total1(12) + Amt
         Total2(12) = Total2(12) + Amt
         Total3(12) = Total3(12) + Amt
         Total4(12) = Total4(12) + Amt
         TempAmount1 = Amt
         
         Amt = Tg.TOTAL_PRICE_RT * Ratio
         Set HeadCf = m_Headers1.Fields(13)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(13) = Total1(13) + Amt
         Total2(13) = Total2(13) + Amt
         Total3(13) = Total3(13) + Amt
         Total4(13) = Total4(13) + Amt
         TempTotalPrice1 = Amt
         
         Amt = 0
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE)
         Amt = TempBd.TOTAL_AMOUNT
         
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(14)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(14) = Total1(14) + Amt
         Total2(14) = Total2(14) + Amt
         Total3(14) = Total3(14) + Amt
         Total4(14) = Total4(14) + Amt
         TempAmount2 = Amt
         
         Amt = 0
         Amt = MyDiff(TempAmount2 * 100, TempAmount1)
         
         Set HeadCf = m_Headers1.Fields(15)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(15) = Total1(15) + Amt
         Total2(15) = Total2(15) + Amt
         Total3(15) = Total3(15) + Amt
         Total4(15) = Total4(15) + Amt
         
         Amt = 0
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE)
         Amt = TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(16) = Total1(16) + Amt
         Total2(16) = Total2(16) + Amt
         Total3(16) = Total3(16) + Amt
         Total4(16) = Total4(16) + Amt
         TempTotalPrice2 = Amt
         
         Amt = 0
         Amt = MyDiff(TempTotalPrice2 * 100, TempTotalPrice1)
         
          Set HeadCf = m_Headers1.Fields(17)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(17) = Total1(17) + Amt
         Total2(17) = Total2(17) + Amt
         Total3(17) = Total3(17) + Amt
         Total4(17) = Total4(17) + Amt
          
          Amt = TempAmount2 - TempAmount1
          
         Set HeadCf = m_Headers1.Fields(18)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(18) = Total1(18) + Amt
         Total2(18) = Total2(18) + Amt
         Total3(18) = Total3(18) + Amt
         Total4(18) = Total4(18) + Amt
         
         Amt = (1 - MyDiff(TempAmount2, TempAmount1)) * 100
         
         Set HeadCf = m_Headers1.Fields(19)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(19) = Total1(19) + Amt
         Total2(19) = Total2(19) + Amt
         Total3(19) = Total3(19) + Amt
         Total4(19) = Total4(19) + Amt
         
         Amt = TempTotalPrice2 - TempTotalPrice1
         
         Set HeadCf = m_Headers1.Fields(20)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(20) = Total1(20) + Amt
         Total2(20) = Total2(20) + Amt
         Total3(20) = Total3(20) + Amt
         Total4(20) = Total4(20) + Amt
         
         Amt = (1 - MyDiff(TempTotalPrice2, TempTotalPrice1)) * 100
         
         Set HeadCf = m_Headers1.Fields(21)
         TempStr = FormatNumberToNull(Amt, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(21) = Total1(21) + 0
         Total2(21) = Total2(21) + 0
         Total3(21) = Total3(21) + 0
         Total4(21) = Total4(21) + 0
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
      Next m_BillingDoc
      
      Call GenerateFooter("      รวม" & PrevKey3, Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("รวม เขตการขาย   : " & PrevKey2 & "", Total3, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      For Each Tg In m_TagetDetails
         If Tg.Flag <> "Y" Then
            Total4(2) = Total4(2) + Tg.TOTAL_AMOUNT * Ratio
            Total4(3) = Total4(3) + Tg.TOTAL_PRICE * Ratio
         End If
      Next Tg

      Call GenerateFooter("      สุทธิ" & "" & "", Total4, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      
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

Private Sub GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

Dim TempAmount1 As Double
Dim TempAmount2 As Double
Dim TempTotalPrice1 As Double
Dim TempTotalPrice2 As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Or j = 12 Then
         TempStr = FormatNumberToNull(Tot(j))
         TempAmount1 = Tot(j)
      ElseIf j = 3 Or j = 13 Then
         TempStr = FormatNumberToNull(Tot(j))
         TempTotalPrice1 = Tot(j)
      ElseIf j = 4 Or j = 14 Then
         TempStr = FormatNumberToNull(Tot(j))
         TempAmount2 = Tot(j)
      ElseIf j = 5 Or j = 15 Then
         TempStr = FormatNumberToNull(MyDiff(TempAmount2 * 100, TempAmount1), 0)
      ElseIf j = 6 Or j = 16 Then
         TempStr = FormatNumberToNull(Tot(j))
         TempTotalPrice2 = Tot(j)
      ElseIf j = 7 Or j = 17 Then
         TempStr = FormatNumberToNull(MyDiff(TempTotalPrice2 * 100, TempTotalPrice1), 0)
      ElseIf j = 8 Or j = 18 Or j = 10 Or j = 20 Or j = 22 Then
         TempStr = FormatNumberToNull(Tot(j))
      ElseIf j = 9 Then
         TempStr = FormatNumberToNull((MyDiff(TempAmount2, TempAmount1) - 1) * 100, 0)
      ElseIf j = 19 Then
         TempStr = FormatNumberToNull((1 - MyDiff(TempAmount2, TempAmount1)) * 100, 0)
      ElseIf j = 11 Then
         TempStr = FormatNumberToNull((MyDiff(TempTotalPrice2, TempTotalPrice1) - 1) * 100, 0)
      ElseIf j = 21 Then
         TempStr = FormatNumberToNull((1 - MyDiff(TempTotalPrice2, TempTotalPrice1)) * 100, 0)
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
   GetKey1 = Trim(Tg.CUSTOMER_BRANCH_CODE & "-" & Tg.SALE_CODE & "-" & Tg.STOCK_NO)
End Function


