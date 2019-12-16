VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling012_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling012_1"
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
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double

Private CreditBalances As Collection
Private TotalMoves As Collection
Private m_Main As Collection
Private m_CheckMain As Collection
Private m_PaidAmounts As Collection
Private DateCount As Long
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   Set m_Main = New Collection
   Set m_CheckMain = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set CreditBalances = New Collection
   Set TotalMoves = New Collection
   Set m_PaidAmounts = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details2 = Nothing
   Set CreditBalances = Nothing
   Set TotalMoves = Nothing
   Set m_Main = Nothing
   Set m_CheckMain = Nothing
   Set m_PaidAmounts = Nothing
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
Dim Amt As Double
Dim iCount As Long
Dim DateCount  As Long
Dim TempDateString As String
Dim TempFromDate  As Date
Dim TempToDate   As Date
Dim TempWidth   As Double

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   Set Cf = New CReportField
   
   TempWidth = 0
   
   Call Cf.SetFieldValue(10, "^", "ลูกค้า", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ครบกำหนด", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ใบกำกับ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันที่", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "ใบวางบิล", "^")
   Call m_Headers1.AddField(Cf)
   
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   
   TempFromDate = mcolParam("FROM_DATE")
   Call Cf.SetFieldValue(2, "^", "เกินกำหนดชำระ", ">")
   Call m_Headers1.AddField(Cf)
   TempWidth = TempWidth + 2
   
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
      TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)

      Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
      Call m_Headers1.AddField(Cf)
      
      TempWidth = TempWidth + 2
      TempFromDate = DateAdd("D", 1, TempToDate)
   Wend
   
   Call Cf.SetFieldValue(2, "^", "หลังวันที่ " & TempToDate, ">")
   Call m_Headers1.AddField(Cf)
   TempWidth = TempWidth + 2
   
   Call Cf.SetFieldValue(TempWidth, "^", "ช่วงวันที่", "<")
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(2, "^", "ยอดค้างรวม", "<")
   Call m_Headers0.AddField(Cf)
   
   
   Set Cf = Nothing
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
Dim TempFromDate As Date
Dim TempToDate As Date
    
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   Vsp.ZoomMode = zmPageWidth
   
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
   
   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = DateAdd("D", DateCount - 1, mcolParam("FROM_DATE"))
   
   Call GetPaidAmountBySaleCode(m_PaidAmounts, , mcolParam("TO_RCP_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), RECEIPT2_DOCTYPE)
   
   Call GetDebtorAging(m_Main, m_CheckMain, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("SHORT_CODE"))
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(TempFromDate), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(TempToDate), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " ถึงวันที่รับชำระ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_RCP_DATE")), "N/A") & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
                                       
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
Dim M As Long
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
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As Double
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim TempFromDate As Date
Dim TempToDate As Date
Dim TempDateString As String
Dim CountDate As Long
Dim TempDate As Date
Dim TempDateCount As Long
Dim Rcp As CRcpCnDn_Item

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
   
   Set m_BillingDoc = New CBillingDoc

   I = 0
   k = 0
   L = 0
   M = 0
   For Each m_BillingDoc In m_Main
      Call m_Details1.ClearField
'      Set Rcp = GetObject("CRcpCnDn_Item", m_PaidAmounts, Trim(Str(m_BillingDoc.BILLING_DOC_ID) & "-" & Str(m_BillingDoc.DOCUMENT_TYPE) & "-" & m_BillingDoc.APAR_CODE))
      Set Rcp = GetObject("CRcpCnDn_Item", m_PaidAmounts, Trim(Str(m_BillingDoc.BILLING_DOC_ID) & "-" & m_BillingDoc.APAR_CODE))
      Set TempBd = GetObject("CBillingDoc", m_CheckMain, Trim(m_BillingDoc.DOCUMENT_NO))
      
      If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
         Amt = TempBd.TOTAL_PRICE + TempBd.VAT_AMOUNT - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT - TempBd.PAID_DISCOUNT - Rcp.PAID_AMOUNT
      ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
         Amt = TempBd.PAY_AMOUNT - Rcp.PAID_AMOUNT - TempBd.PAID_DISCOUNT
      End If

      If Round(Amt, 2) <> 0 Then

      Amt = 0
      If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
         Amt = m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT - m_BillingDoc.PAID_DISCOUNT
         If m_BillingDoc.VAT_AMOUNT <> 0 Then
            Amt = Amt + (Amt * 0.07)
         End If
          If Rcp.PAID_AMOUNT <> 0 Then
            Amt = Amt - Rcp.PAID_AMOUNT
         End If
      ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
         Amt = m_BillingDoc.PAY_AMOUNT - Rcp.PAID_AMOUNT - m_BillingDoc.PAID_DISCOUNT
      End If
      If Round(Amt, 2) <> 0 Then
         If PrevKey6 <> m_BillingDoc.SALE_CODE And I <> 0 Then
            If mcolParam("SHOW_DETAIL") = 1 Then
               Call GenerateFooter("", "     รวมลูกค้า " & PrevKey4 & " สาขา " & PrevKey7, I & " ใบ", Total1, m_Details1, 0)
            Else
               Call GenerateFooter("", PrevKey4 & " สาขา " & PrevKey7 & " (" & PrevKey3 & ")", I & " ใบ", Total1, m_Details1, 0)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
         
            Call GenerateFooter("", "รวมพนักงานขาย " & PrevKey6, M & " ใบ", Total3, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            I = 0
            M = 0
            Vsp.FontBold = True
            Vsp = "พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            Vsp.FontBold = False
            
            If mcolParam("SHOW_DETAIL") = 1 Then
               Vsp.FontBold = True
               Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_LONG_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
               Vsp.FontBold = False
            End If
            
            For j = 1 To UBound(Total3)
               Total1(j) = 0
               Total3(j) = 0
            Next j
         ElseIf PrevKey5 <> m_BillingDoc.CUSTOMER_BRANCH And I <> 0 Then
            If mcolParam("SHOW_DETAIL") = 1 Then
               Call GenerateFooter("", "     รวมลูกค้า " & PrevKey4 & " สาขา " & PrevKey7, I & " ใบ", Total1, m_Details1, 0)
            Else
               Call GenerateFooter("", PrevKey4 & " สาขา " & PrevKey7 & " (" & PrevKey3 & ")", I & " ใบ", Total1, m_Details1, 0)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            I = 0

            If mcolParam("SHOW_DETAIL") = 1 Then
               Vsp.FontBold = True
               Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_LONG_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
               Vsp.FontBold = False
            End If
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         ElseIf I = 0 Then
            Vsp.FontBold = True
            Vsp = "พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            If mcolParam("SHOW_DETAIL") = 1 Then
               Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_LONG_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
            End If
            Vsp.FontBold = False
         End If

         PrevKey3 = m_BillingDoc.APAR_CODE
         PrevKey5 = m_BillingDoc.CUSTOMER_BRANCH
         PrevKey4 = m_BillingDoc.APAR_LONG_NAME
         PrevKey6 = m_BillingDoc.SALE_CODE
         PrevKey7 = m_BillingDoc.CUSTOMER_BRANCH_NAME

         I = I + 1
         L = L + 1
         M = M + 1
         SumRow = 0
         Amt = 0
         Call m_Details1.ClearField
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = DateToStringExtEx2(m_BillingDoc.Due_Date)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = m_BillingDoc.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
'         TempStr = m_BillingDoc.DOC_ID_BILLS_NO
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
            If m_BillingDoc.VAT_AMOUNT <> 0 Then
               Amt = Amt + (Amt * 0.07)
            End If
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = -(m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT)
            If m_BillingDoc.VAT_AMOUNT <> 0 Then
               Amt = Amt + (Amt * 0.07)
            End If
         ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Then
            Amt = -m_BillingDoc.PAY_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
            Amt = m_BillingDoc.PAY_AMOUNT
         End If
         SumRow = Amt
         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = Rcp.PAID_AMOUNT + m_BillingDoc.PAID_DISCOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = -(Rcp.PAID_AMOUNT + m_BillingDoc.PAID_DISCOUNT)
         ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Then
            Amt = -(Rcp.PAID_AMOUNT + m_BillingDoc.PAID_DISCOUNT)
         ElseIf m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
            Amt = Rcp.PAID_AMOUNT + m_BillingDoc.PAID_DISCOUNT
         End If
         SumRow = SumRow - Amt

         k = 4
         TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
         TempDateString = Trim(Replace(TempDateString, ")", ""))
         
         TempFromDate = mcolParam("FROM_DATE")
      
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         If m_BillingDoc.Due_Date < TempFromDate Then
            TempStr = FormatNumberToNull(SumRow)
            Total1(k) = Total1(k) + SumRow
            Total2(k) = Total2(k) + SumRow
            Total3(k) = Total3(k) + SumRow
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         DateCount = 0
         While Len(TempDateString) > 0
               Call CalculateDatePeriod(TempDateString, DateCount)
               TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
               TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)
                  
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               If TempFromDate <= m_BillingDoc.Due_Date And m_BillingDoc.Due_Date <= TempToDate Then
                  TempStr = FormatNumberToNull(SumRow)
                  Total1(k) = Total1(k) + SumRow
                  Total2(k) = Total2(k) + SumRow
                  Total3(k) = Total3(k) + SumRow
               Else
                  TempStr = ""
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)

               TempFromDate = DateAdd("D", 1, TempToDate)
         Wend

         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         If m_BillingDoc.Due_Date > TempToDate Then
            TempStr = FormatNumberToNull(SumRow)
            Total1(k) = Total1(k) + SumRow
            Total2(k) = Total2(k) + SumRow
            Total3(k) = Total3(k) + SumRow
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = ""
         Total1(k) = Total1(k) + SumRow
         Total2(k) = Total2(k) + SumRow
         Total3(k) = Total3(k) + SumRow
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         If mcolParam("SHOW_DETAIL") = 1 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      End If

      End If
   Next m_BillingDoc

   If mcolParam("SHOW_DETAIL") = 1 Then
      Call GenerateFooter("", "     รวมลูกค้า " & PrevKey4 & " สาขา " & PrevKey7, I & " ใบ", Total1, m_Details1, 0)
   Else
      Call GenerateFooter("", PrevKey4 & "สาขา " & PrevKey7 & " (" & PrevKey3 & ")", I & " ใบ", Total1, m_Details1, 0)
   End If
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("", "รวมพนักงานขาย " & PrevKey6, M & " ใบ", Total3, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

   Call GenerateFooter("", "สุทธิ", L & " ใบ", Total2, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set m_BillingDoc = Nothing
   
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
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   For j = 1 To 3
      Set HeadCf = m_Headers1.Fields(j)
      TempWidth = TempWidth + HeadCf.Width
   Next j
   TempStr = txt1 & " " & txt2
   Call BodyCf.SetFieldValue(TempWidth, "<", TempStr, "<")
   Call Details.AddField(BodyCf)
   
   Set HeadCf = m_Headers1.Fields(4)
   TempStr = Txt3
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j >= 5 And j <= 13 Then
         TempStr = FormatNumberToNull(Tot(j))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      ElseIf j > 13 Then
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
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
Private Sub GenerateHeader(txt1 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
Private Sub CalculateDate(TempStr As String, DateCount As Long)
Dim TempID As Long
Dim TempStrNew  As String
   TempStrNew = Replace(TempStr, "(", "")
   TempStrNew = Replace(TempStrNew, ")", "")
   TempID = InStr(1, TempStrNew, ",")
   While InStr(1, TempStrNew, ",") > 0
      TempID = InStr(1, TempStrNew, ",")
      DateCount = DateCount + Val(Left(TempStrNew, TempID - 1))
      TempStrNew = Mid(TempStrNew, TempID + 1)
   Wend
   DateCount = DateCount + Val(TempStrNew)
End Sub
Private Sub CalculateDatePeriod(TempStr As String, DateCount As Long)
Dim TempID As Long
   TempID = InStr(1, TempStr, ",")
   If TempID > 0 Then
      DateCount = Val(Left(TempStr, TempID - 1))
      TempStr = Mid(TempStr, TempID + 1)
   Else
      DateCount = Val(TempStr)
      TempStr = ""
   End If
End Sub
