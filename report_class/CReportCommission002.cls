VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCommission002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCommission002"
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
Private TempDateCountTaget As Long
Private TempFromDate  As Date
Private TempToDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<52|<4|<8;"

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
Private m_TempFirstDistinctMain As Collection
Private m_SaleAmounts As Collection
Private m_SaleAmountExs As Collection

Private m_TagetDetails As Collection
Private m_TagetDetailsSaleCode2 As Collection
Private SumTotalChart As Collection
Private SaleChartColl As Collection
Private OrderSaleChartColl As Collection
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
   Set m_TempFirstDistinctMain = New Collection
   Set m_TagetDetails = New Collection
   Set m_TagetDetailsSaleCode2 = New Collection
   Set SumTotalChart = New Collection
   Set SaleChartColl = New Collection
   Set OrderSaleChartColl = New Collection
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
   Set m_TempFirstDistinctMain = Nothing
   Set m_TagetDetails = Nothing
   Set m_TagetDetailsSaleCode2 = Nothing
   Set SumTotalChart = Nothing
   Set SaleChartColl = Nothing
   Set OrderSaleChartColl = Nothing
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
Dim TempFromDate  As Date
Dim TempToDate   As Date

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

   Call Cf.SetFieldValue(4, "^", "รหัสพนักงานขาย / พนักงานขาย", "<")
   Call m_Headers1.AddField(Cf)

   'สะสม
'   Call Cf.SetFieldValue(8, "^", "สะสม " & DateToStringExtEx2(mcolParam("FROM_DATE")) & "-" & DateToStringExtEx2(mcolParam("TO_DATE")), ">")
'   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เป้าการขาย", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Diff", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing

'   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
'   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'   strPrint = TempStr1_1
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
Dim D As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 500
   Vsp.MarginFooter = 500
   Vsp.MarginHeader = 500
   Vsp.MarginLeft = 500
   Vsp.MarginRight = 500
   Vsp.MarginTop = 500
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   Vsp.ZoomMode = zmPageWidth
   
'   Call GetFirstLastDate(mcolParam("FROM_DATE"), TempFromDate, TempToDate)
   TempDateCountTaget = DateDiff("D", mcolParam("FROM_DATE"), mcolParam("TO_DATE")) + 1 - Val(mcolParam("HOLIDAY"))
   
   Call GetDistinctEmpDocTypeFree(m_TempFirstDistinctMain, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   Call GetSaleAmountEmpDocType(m_SaleAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   Call GetSaleAmountEmpDocTypeBranch(m_SaleAmountExs, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))

   Call LoadTagetDetailEmp(m_TagetDetails, Trim((Val(Year(mcolParam("FROM_DATE")))) & Format(Month(mcolParam("FROM_DATE")), "00")))
   Call LoadTagetDetailSaleCode2(m_TagetDetailsSaleCode2, Trim((Val(Year(mcolParam("FROM_DATE")))) & Format(Month(mcolParam("FROM_DATE")), "00")))
   
   Call SumTempDistinctMain2(m_TempFirstDistinctMain, m_TagetDetailsSaleCode2, m_DistinctMain)
   
   FromDate = "จากงวด " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " " & "จากพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A") & _
                                       "จากสินค้า" & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A")
                                       
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
Dim TempBd As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim TempDate As Date
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
   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True

      Set BodyCf = New CReportField
      I = 0
      
      Dim m_SaleChart As CSaleChart
      Dim TotalSale As CTotalSale
      
      Set m_SaleChart = New CSaleChart
      Set TotalSale = New CTotalSale
      
      m_SaleChart.SALE_CHART_ID = -1
      m_SaleChart.FROM_DATE = mcolParam("FROM_DATE")
      m_SaleChart.TO_DATE = mcolParam("TO_DATE")
      Call m_SaleChart.QueryData(2, Rs, iCount)

      If iCount <= 0 Then
         mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
         Exit Function
      End If

      While Not Rs.EOF
         Set m_SaleChart = New CSaleChart
         Call m_SaleChart.PopulateFromRS(2, Rs)
         Call SaleChartColl.add(m_SaleChart, Trim(Str(m_SaleChart.SALE_CHART_ID)))
         Rs.MoveNext
      Wend
      
      Call GenerateOrderSaleChart(SaleChartColl, -1, 0)
      
      Call SumChart(OrderSaleChartColl, SumTotalChart)
      
      For Each m_BillingDoc In m_DistinctMain
         I = I + 1
         Amt = 0
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_BillingDoc.SALE_CODE & " /  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set Tg = GetObject("CTagetDetail", m_TagetDetails, m_BillingDoc.SALE_CODE)
         'เป้าสะสม
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = FormatNumberToNull(Tg.TOTAL_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(2) = Total1(2) + Tg.TOTAL_PRICE
         'ยอดสะสม
         
         Amt = 0
         Set HeadCf = m_Headers1.Fields(3)
         
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         If mcolParam("NOT_SHOW_RETURN") <> "1" Then
            Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE)
            Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
         End If
            
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         If mcolParam("NOT_SHOW_RETURN") <> "1" Then
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE)
            Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
         End If
         
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(3) = Total1(3) + Amt
         'Diff สะสม
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumberToNull(Tg.TOTAL_PRICE - Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(4) = Total1(4) + (Tg.TOTAL_PRICE - Amt)
         '% สะสม
         Set HeadCf = m_Headers1.Fields(5)
         If Tg.TOTAL_PRICE > 0 Then
            TempStr = FormatNumberToNull((Amt * 100) / (Tg.TOTAL_PRICE))
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Next m_BillingDoc
      
      If Total1(2) <> 0 Then
         Total1(5) = ((Total1(3) * 100) / Total1(2))
      End If
      
      Call GenerateFooter("      สุทธิ" & "" & "", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Vsp = ""
      
      Call printHeader
      
      For Each m_SaleChart In OrderSaleChartColl
      
         Call m_Details1.ClearField
   
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_SaleChart.SALE_NAME & "(" & m_SaleChart.SALE_CODE & ")"
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set TotalSale = GetObject("CTotalSale", SumTotalChart, Trim(Str(m_SaleChart.EMP_ID)))
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = FormatNumberToNull(TotalSale.TAGET_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
               
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = FormatNumberToNull(TotalSale.TOTAL_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumberToNull(TotalSale.TAGET_PRICE - TotalSale.TOTAL_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(5)
         If TotalSale.TAGET_PRICE > 0 Then
            TempStr = FormatNumberToNull((TotalSale.TOTAL_PRICE * 100) / (TotalSale.TAGET_PRICE))
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If mcolParam("SHOW_COLOR") = 1 Then
            If m_SaleChart.Level = 0 Then
               Vsp.TextColor = RGB(0, 0, 0)
            ElseIf m_SaleChart.Level = 1 Then
               Vsp.TextColor = RGB(255, 0, 0)
            ElseIf m_SaleChart.Level = 2 Then
               Vsp.TextColor = RGB(0, 0, 255)
            ElseIf m_SaleChart.Level = 3 Then
               Vsp.TextColor = RGB(0, 255, 0)
            Else
               Vsp.TextColor = RGB(((m_SaleChart.Level + 1) * 200) Mod 256, (m_SaleChart.Level * 50) Mod 256, (m_SaleChart.Level * 90) Mod 256)
            End If
         End If
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Next m_SaleChart
      
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

Private Sub GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
   Vsp.TextAlign = taLeftBottom
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
   GetKey1 = Trim(Tg.SALE_CODE & "-" & Tg.APAR_CODE & "-" & Tg.CUSTOMER_BRANCH & "-" & Tg.STOCK_NO)
End Function
Private Sub GenerateEndDoc()
End Sub
Private Sub SumTempDistinctMain2(Cl As Collection, Cl2 As Collection, Cl3 As Collection)
Dim m_BillingDoc As CBillingDoc
Dim m_BillingDoc2 As CBillingDoc
Dim m_TagetDetail As CTagetDetail
Dim TempData As CBillingDoc
Dim TempTagetDetail As CTagetDetail
Dim TempKey1 As String

   Set Cl3 = Nothing
   Set Cl3 = New Collection
   TempKey1 = ""
   For Each m_BillingDoc In Cl
      If TempKey1 <> m_BillingDoc.SALE_CODE Then
         For Each m_TagetDetail In Cl2
            If m_BillingDoc.SALE_CODE > m_TagetDetail.EMPLOYEE_CODE And m_TagetDetail.Flag <> "Y" Then
               Set m_BillingDoc2 = New CBillingDoc
               m_BillingDoc2.SALE_CODE = m_TagetDetail.EMPLOYEE_CODE
               m_BillingDoc2.SALE_JOINT_CODE = m_TagetDetail.SALE_JOINT_CODE
               m_BillingDoc2.SALE_LONG_NAME = m_TagetDetail.SALE_LONG_NAME
               m_BillingDoc2.SALE_LAST_NAME = m_TagetDetail.SALE_LAST_NAME
               m_TagetDetail.Flag = "Y"
               Call Cl3.add(m_BillingDoc2)
            End If
         Next m_TagetDetail
      End If

      For Each m_TagetDetail In Cl2
         If m_BillingDoc.SALE_CODE = m_TagetDetail.EMPLOYEE_CODE And m_TagetDetail.Flag <> "Y" Then
            m_TagetDetail.Flag = "Y"
         End If
      Next m_TagetDetail
      Call Cl3.add(m_BillingDoc)
      TempKey1 = m_BillingDoc.SALE_CODE
   Next m_BillingDoc
   'เติมที่เหลือ
   For Each m_TagetDetail In Cl2
      If m_TagetDetail.Flag <> "Y" Then
         Set m_BillingDoc2 = New CBillingDoc
         m_BillingDoc2.SALE_CODE = m_TagetDetail.EMPLOYEE_CODE
         m_BillingDoc2.SALE_JOINT_CODE = m_TagetDetail.SALE_JOINT_CODE
         m_BillingDoc2.SALE_LONG_NAME = m_TagetDetail.SALE_LONG_NAME
         m_BillingDoc2.SALE_LAST_NAME = m_TagetDetail.SALE_LAST_NAME
         Call Cl3.add(m_BillingDoc2)
      End If
   Next m_TagetDetail
End Sub
Private Sub SumChart(Coll As Collection, SumTotal As Collection)
Dim Cm As CSaleChart
   For Each Cm In Coll
      Call Recuresive(Cm, SumTotal, GetParent(Cm.SALE_CHART_ID), GetEmp(Cm.SALE_CHART_ID))
   Next Cm
End Sub
Public Sub Recuresive(Cm As CSaleChart, SumTotal As Collection, ParentID As Long, OwnId As Long)
On Error Resume Next
Dim Amt As Double
Dim P1 As CTotalSale
Dim P2 As CTotalSale
Set P1 = New CTotalSale
Dim Old As Double
Dim TempBd As CBillingDoc
Dim Tg As CTagetDetail


   P1.EMP_ID = OwnId
   P1.SALE_NAME = Cm.SALE_NAME & " (" & Cm.SALE_CODE & ")"
   
   Amt = 0
         
   Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Cm.SALE_CODE & "-" & INVOICE_DOCTYPE)
   Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
   Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Cm.SALE_CODE & "-" & RECEIPT1_DOCTYPE)
   Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
   If mcolParam("NOT_SHOW_RETURN") <> "1" Then
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Cm.SALE_CODE & "-" & RETURN_DOCTYPE)
      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
   End If
      
   Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, Cm.SALE_CODE & "-" & INVOICE_DOCTYPE)
   Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
   Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, Cm.SALE_CODE & "-" & RECEIPT1_DOCTYPE)
   Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
   If mcolParam("NOT_SHOW_RETURN") <> "1" Then
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, Cm.SALE_CODE & "-" & RETURN_DOCTYPE)
      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
   End If
         
   P1.TOTAL_PRICE = Amt
   
   Set Tg = GetObject("CTagetDetail", m_TagetDetails, Cm.SALE_CODE)
   P1.TAGET_PRICE = Tg.TOTAL_PRICE
   
         
   If SumTotal.Count = 0 Then
      Call SumTotal.add(P1, Trim(P1.Getkey))
   Else
      Set P2 = SumTotal(Trim(P1.Getkey))
      If P2 Is Nothing Then
         Call SumTotal.add(P1, Trim(P1.Getkey))
      Else
         P2.TAGET_PRICE = P2.TAGET_PRICE + P1.TAGET_PRICE
         P2.TOTAL_PRICE = P2.TOTAL_PRICE + P1.TOTAL_PRICE
      End If
   End If
   
   If ParentID > 0 Then
      Call Recuresive(Cm, SumTotal, GetParent(ParentID), GetEmp(ParentID))
   End If
End Sub
Private Function GetParent(ID As Long) As Long
Dim Cm As CSaleChart
   Set Cm = GetObject("CSaleChart", SaleChartColl, Trim(Str(ID)))
   GetParent = Cm.PARENT_ID
End Function
Private Function GetEmp(ID As Long) As Long
Dim Cm As CSaleChart
   Set Cm = GetObject("CSaleChart", SaleChartColl, Trim(Str(ID)))
   GetEmp = Cm.EMP_ID
End Function
Private Sub GenerateOrderSaleChart(TempColl As Collection, PID As Long, Level As Long)
Dim O As CSaleChart

   For Each O In TempColl
      If O.PARENT_ID = PID Then
         O.SALE_NAME = Space(Level * 5) & O.SALE_NAME
         O.Level = Level
         Call OrderSaleChartColl.add(O, Trim(Str(O.SALE_CHART_ID)))
         Call GenerateOrderSaleChart(TempColl, O.SALE_CHART_ID, Level + 1)
      End If
   Next O
End Sub

