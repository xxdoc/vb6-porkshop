VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling029"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling024"
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

Private m_AddColls As Collection
Private m_SumAddColls As Collection
Private m_SubColls As Collection
Private m_SumSubColls As Collection
Private m_AccountColls As Collection
Private m_SumAccountColls As Collection

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
   
   Set m_AddColls = New Collection
   Set m_SumAddColls = New Collection
   Set m_SubColls = New Collection
   Set m_SumSubColls = New Collection
   Set m_AccountColls = New Collection
   Set m_SumAccountColls = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_SubColls = Nothing
   Set m_AddColls = Nothing
   Set m_SumAddColls = Nothing
   Set m_SumSubColls = Nothing
   Set m_AccountColls = Nothing
   Set m_SumAccountColls = Nothing
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
Dim iCount As Double
Dim BdAdd As CBillingAddition
Dim BdSub As CBillingSubTract
Dim Ct As CCashTran
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   iCount = 0
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(3, "^", "เลขที่", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ว.รับเงิน", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดหนี้", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดชำระ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ส่วนเพิ่ม", ">")
   Call m_Headers1.AddField(Cf)
   iCount = iCount + 2
      
   If mcolParam("SHOW_DETAIL_ADDSUB") = 1 Then
      For Each BdAdd In m_AddColls
         Call Cf.SetFieldValue(2, "^", BdAdd.ADDITION_CODE, ">")
         Call m_Headers1.AddField(Cf)
         iCount = iCount + 2
      Next BdAdd
   End If
   
   Call Cf.SetFieldValue(2, "^", "ส่วนหัก", ">")
   Call m_Headers1.AddField(Cf)
   iCount = iCount + 2
   
   If mcolParam("SHOW_DETAIL_ADDSUB") = 1 Then
      For Each BdSub In m_SubColls
         Call Cf.SetFieldValue(2, "^", BdSub.SUBTRACT_CODE, ">")
         Call m_Headers1.AddField(Cf)
         iCount = iCount + 2
      Next BdSub
   End If
   
   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers1.AddField(Cf)
   iCount = iCount + 2
   
   Call Cf.SetFieldValue(2, "^", "VAT", ">")
   Call m_Headers1.AddField(Cf)
   iCount = iCount + 2
   
   Call Cf.SetFieldValue(2, "^", "สุทธิ", ">")
   Call m_Headers1.AddField(Cf)
   iCount = iCount + 2
   If mcolParam("SHOW_PAYMENT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "เงินสด", ">")
      Call m_Headers1.AddField(Cf)
      iCount = iCount + 2
      
      Call Cf.SetFieldValue(2, "^", "เช็ค", ">")
      Call m_Headers1.AddField(Cf)
      iCount = iCount + 2
      
      If mcolParam("SHOW_ACCOUNT") = 1 Then
         For Each Ct In m_AccountColls
            Call Cf.SetFieldValue(2, "^", Ct.BANK_ACCOUNT_NAME, ">")
            Call m_Headers1.AddField(Cf)
            iCount = iCount + 2
         Next Ct
      Else
         Call Cf.SetFieldValue(2, "^", "โอน", ">")
         Call m_Headers1.AddField(Cf)
         iCount = iCount + 2
      End If
      
      Call Cf.SetFieldValue(2, "^", "FEE", ">")
      Call m_Headers1.AddField(Cf)
      iCount = iCount + 2
      
      Call Cf.SetFieldValue(2, "^", "ส่วนต่าง", ">")
      Call m_Headers1.AddField(Cf)
      iCount = iCount + 2
      
   End If
   
   Call Cf.SetFieldValue(3, "^", "เลขที่", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "วันที่", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "จำนวนเงิน", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(iCount, "^", "", "^")
   Call m_Headers0.AddField(Cf)
   
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
   
   If mcolParam("SHOW_DETAIL_ADDSUB") = 1 Then
      Call GetDistinctBillingAddition(m_AddColls, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
      Call GetSumBillingAdditionID(m_SumAddColls, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
      Call GetDistinctBillingSubTract(m_SubColls, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
      Call GetSumBillingSubTractID(m_SumSubColls, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
   End If
   If mcolParam("SHOW_ACCOUNT") = 1 Then
      Call GetDistinctBillBankAccount(m_AccountColls, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
      Call GetTransferAmountBillBankAccount(m_SumAccountColls, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"))
   End If
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากรหัสพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงรหัสพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
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
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Sum1 As Double
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim k As Long
Dim BdAdd As CBillingAddition
Dim TempAdd As CBillingAddition
Dim BdSub As CBillingSubTract
Dim TempSub As CBillingSubTract
Dim Ct As CCashTran
Dim TempCt As CCashTran

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
      
      MasterInd = "73"
      Set m_BillingDoc = New CBillingDoc
      MasterInd = "1"
      
      m_BillingDoc.BILLING_DOC_ID = -1
      m_BillingDoc.DOCUMENT_TYPE = RECEIPT2_DOCTYPE
      m_BillingDoc.APAR_IND = 1
      m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
      m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
      m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
      m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
      m_BillingDoc.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      m_BillingDoc.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      Call m_BillingDoc.QueryData(73, Rs, iCount)
         
      I = 0
      If mcolParam("PRINT_TO_FILE") = 1 Then
         Call printHeader
      End If
      While Not Rs.EOF
         I = I + 1
         Call m_BillingDoc.PopulateFromRS(73, Rs)
         
         Amt = 0
         
         Call m_Details1.ClearField
         
         If PrevKey1 <> m_BillingDoc.SALE_CODE And I <> 1 Then
            Call GenerateFooter("      รวม " & PrevKey1, Total1, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
   
            Vsp = "พนักงานขาย  :  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            ' รวม
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         ElseIf I = 1 Then
            Vsp = "พนักงานขาย  :  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
         End If
         
         PrevKey1 = m_BillingDoc.SALE_CODE
         PrevKey2 = m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
         
         If PrevKey3 <> m_BillingDoc.DOCUMENT_NO Then
            
            PrevKey3 = m_BillingDoc.DOCUMENT_NO
            
            Call m_Details1.ClearField
         
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = m_BillingDoc.DOCUMENT_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = DateToStringExtEx2(m_BillingDoc.Due_Date)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(4)
            TempStr = m_BillingDoc.APAR_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(5)
            Amt = m_BillingDoc.PAY_AMOUNT + m_BillingDoc.DEBIT_AMOUNT - m_BillingDoc.CREDIT_AMOUNT
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(5) = Total1(5) + Amt
            Total2(5) = Total2(5) + Amt
            
            Set HeadCf = m_Headers1.Fields(6)
            Amt = m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT - m_BillingDoc.CREDIT_AMOUNT
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + Amt
            Total2(6) = Total2(6) + Amt
            Sum1 = Amt
            
            k = 6
            
            Amt = m_BillingDoc.ADDITION_AMOUNT
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
            Sum1 = Sum1 + Amt
            
            If mcolParam("SHOW_DETAIL_ADDSUB") = 1 Then
               For Each BdAdd In m_AddColls
                  Set TempAdd = GetObject("CBillingAddition", m_SumAddColls, Trim(m_BillingDoc.BILLING_DOC_ID & "-" & BdAdd.ADDITION_ID))
                  Amt = TempAdd.ITEM_AMOUNT
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(Amt)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Amt
                  Total2(k) = Total2(k) + Amt
               Next BdAdd
            End If
            
            Amt = m_BillingDoc.SUBTRACT_AMOUNT
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
            Sum1 = Sum1 - Amt
            
            If mcolParam("SHOW_DETAIL_ADDSUB") = 1 Then
               For Each BdSub In m_SubColls
                  Set TempSub = GetObject("CBillingSubTract", m_SumSubColls, Trim(m_BillingDoc.BILLING_DOC_ID & "-" & BdSub.SUBTRACT_ID))
                  Amt = TempSub.ITEM_AMOUNT
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(Amt)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Amt
                  Total2(k) = Total2(k) + Amt
               Next BdSub
            End If
            
            Amt = Sum1
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
                
            Amt = m_BillingDoc.VAT_AMOUNT
             k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
            Sum1 = Sum1 + Amt
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            Amt = Sum1
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
            
            If mcolParam("SHOW_PAYMENT") = 1 Then
               Amt = m_BillingDoc.CASH_PMT
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               
               Amt = m_BillingDoc.CHEQUE_PMT
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               
               If mcolParam("SHOW_ACCOUNT") = 1 Then
                  For Each Ct In m_AccountColls
                     Set TempCt = GetObject("CCashTran", m_SumAccountColls, Trim(m_BillingDoc.BILLING_DOC_ID & "-" & Ct.BANK_ACCOUNT))
                     Amt = TempCt.NET_AMOUNT
                     k = k + 1
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                  Next Ct
               Else
                  Amt = m_BillingDoc.BANKTRF_PMT - m_BillingDoc.FEE_AMOUNT
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(Amt)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Amt
                  Total2(k) = Total2(k) + Amt
               End If
               
               Amt = m_BillingDoc.FEE_AMOUNT
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               
               Amt = m_BillingDoc.RCP_DIF
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(Amt)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
            End If
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         If mcolParam("SHOW_DETAIL") = 1 Then
            Call m_Details1.ClearField
            
            Set HeadCf = m_Headers0.Fields(1)
            TempStr = m_BillingDoc.DOC_ID_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(2)
            TempStr = DateToStringExtEx2(m_BillingDoc.DOC_ID_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(3)
            If m_BillingDoc.DOCUMENT_TYPE_RCP = INVOICE_DOCTYPE Then
               TempStr = FormatNumberToNull(m_BillingDoc.DOC_ID_PAID)
            ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = RETURN_DOCTYPE Then
               TempStr = FormatNumberToNull(-1 * m_BillingDoc.DOC_ID_PAID)
            ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = CN_DOCTYPE Then
               TempStr = FormatNumberToNull(-1 * m_BillingDoc.CNDN_AMOUNT)
            ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = DN_DOCTYPE Then
               TempStr = FormatNumberToNull(m_BillingDoc.CNDN_AMOUNT)
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
               
            Set HeadCf = m_Headers0.Fields(4)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(5)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(6)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
               
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Rs.MoveNext
      Wend
      
      Call GenerateFooter("      รวม " & PrevKey1, Total1, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("สุทธิ", Total2, m_Details1)
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
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function
Private Sub GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   For j = 1 To 4
      Set HeadCf = m_Headers1.Fields(j)
      TempWidth = TempWidth + HeadCf.Width
   Next j
   TempStr = txt1
   Call BodyCf.SetFieldValue(TempWidth, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j > 4 Then
         TempStr = FormatNumberToNull(Tot(j))
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

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
