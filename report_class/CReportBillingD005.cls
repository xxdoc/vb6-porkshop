VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBillingD005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBillingD005"
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
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double

Private CreditBalances As Collection
Private TotalMoves As Collection

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
   
   Set CreditBalances = New Collection
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
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set CreditBalances = Nothing
   Set TotalMoves = Nothing
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

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(2, "^", "วันที่", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เลขที่", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เก็บเงิน", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลดหนี้/รับคืน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เพิ่มหนี้", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับชำระ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เงินสดรับ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอนเข้า", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เช็ครับ", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เลขที่เช็ค", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ลงวันที่", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันที่เก็บเงิน", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เลขที่ใบกำกับ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "วันที่", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดชำระ", ">")
   Call m_Headers1.AddField(Cf)
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.TableBorder = tbNone
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
   
   Vsp.TableBorder = tbNone
End Sub
Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If mcolParam("FROM_BILL_DATE") > 0 Then
      Call GetCreditBalance(CreditBalances, -1, DateAdd("D", -1, mcolParam("FROM_BILL_DATE")), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"))
   End If
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A")
                                       
                                       
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
Dim Total1(20) As Double
Dim Total2(20) As Double
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim TotalMove As CBillingDoc
Dim TempChequeNo As String
Dim TempChequeDate As String
Dim TempChequeEffectdate As String

Dim TempDocNo As String
Dim TempDocDate As String
Dim TempDocPaid As String

Dim TempCheque As String
Dim ChequeFlag As Boolean
Dim TempDoc As String
Dim DocFlag As Boolean

Dim TempColl As Collection
Dim BD As CBillingDoc
   
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
   Next j
   
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   
   m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
   m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & "," & RECEIPT2_DOCTYPE & ")"
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   
    Call m_BillingDoc.QueryData(28, Rs, iCount)
      
   I = 0
   
   Vsp.TableBorder = tbNone
   While Not Rs.EOF
      I = I + 1
      Call m_Details1.ClearField
      Call m_BillingDoc.PopulateFromRS(28, Rs)
      
      If PrevKey1 <> m_BillingDoc.APAR_CODE And I <> 1 Then
         
         If I <> 1 And Len(TempChequeNo) > 0 Then
            Call GenerateCheque(TempChequeNo, TempChequeDate, TempChequeEffectdate, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            
            TempChequeNo = ""
            TempChequeDate = ""
            TempChequeEffectdate = ""
         End If
         
         If I <> 1 And Len(TempDocNo) > 0 Then
            Call GenerateDoc(TempDocNo, TempDocDate, TempDocPaid, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            TempDocNo = ""
            TempDocDate = ""
            TempDocPaid = ""
            Set TempColl = Nothing
         End If
         
         Call GenerateFooter("", "รวม " & PrevKey2, Total1, m_Details1, 0)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Vsp.TableBorder = tbAll
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         Vsp.TableBorder = tbNone
         
         Vsp = ""
         Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.APAR_CODE & " )"
         ' รวม
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
         
         Amt = 0
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & INVOICE_DOCTYPE, True)
         Amt = Amt + TotalMove.TOTAL_PRICE + TotalMove.VAT_AMOUNT - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & RETURN_DOCTYPE, True)
         Amt = Amt - (TotalMove.TOTAL_PRICE + TotalMove.VAT_AMOUNT - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT)
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & CN_DOCTYPE, True)
         Amt = Amt - (TotalMove.PAID_AMOUNT)
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & DN_DOCTYPE, True)
         Amt = Amt + (TotalMove.PAID_AMOUNT)
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & RECEIPT2_DOCTYPE, True)
         Amt = Amt - (TotalMove.PAID_AMOUNT - TotalMove.CREDIT_AMOUNT + TotalMove.DEBIT_AMOUNT)
         
         Call GenerateBalance("ยกมา", Amt, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         SumRow = Amt
         
      ElseIf I = 1 Then
         Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.APAR_CODE & " )"
         
         Amt = 0
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & INVOICE_DOCTYPE, True)
         Amt = Amt + TotalMove.TOTAL_PRICE + TotalMove.VAT_AMOUNT - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & RETURN_DOCTYPE, True)
         Amt = Amt - (TotalMove.TOTAL_PRICE + TotalMove.VAT_AMOUNT - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT)
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & CN_DOCTYPE, True)
         Amt = Amt - (TotalMove.PAID_AMOUNT)
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & DN_DOCTYPE, True)
         Amt = Amt + (TotalMove.PAID_AMOUNT)
         Set TotalMove = GetObject("CBillingDoc", CreditBalances, m_BillingDoc.APAR_MAS_ID & "-" & RECEIPT2_DOCTYPE, True)
         Amt = Amt - (TotalMove.PAID_AMOUNT - TotalMove.CREDIT_AMOUNT + TotalMove.DEBIT_AMOUNT)
         
         Call GenerateBalance("ยกมา", Amt, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         SumRow = Amt
         
      End If
      
      
      If PrevKey3 <> m_BillingDoc.DOCUMENT_NO Or I = 1 Then
      
         If I <> 1 And Len(TempChequeNo) > 0 Then
            Call GenerateCheque(TempChequeNo, TempChequeDate, TempChequeEffectdate, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            
            TempChequeNo = ""
            TempChequeDate = ""
            TempChequeEffectdate = ""
         End If
         
         If I <> 1 And Len(TempDocNo) > 0 Then
            Call GenerateDoc(TempDocNo, TempDocDate, TempDocPaid, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            TempDocNo = ""
            TempDocDate = ""
            TempDocPaid = ""
            Set TempColl = Nothing
         End If
         
         PrevKey1 = m_BillingDoc.APAR_CODE
         PrevKey2 = m_BillingDoc.APAR_NAME
         
         Call m_Details1.ClearField
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = m_BillingDoc.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         If m_BillingDoc.DOCUMENT_TYPE = RECEIPT2_DOCTYPE Then
            TempStr = ""
         Else
            If m_BillingDoc.DOC_ID1_PAID > 0 Then
               TempStr = "Y"
            Else
               TempStr = "N"
            End If
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
            TempStr = FormatNumberToNull(Amt)
            Total1(4) = Total1(4) + Amt
            Total2(4) = Total2(4) + Amt
            
            SumRow = SumRow + (Amt)
         Else
            Amt = 0
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
   
         Set HeadCf = m_Headers1.Fields(5)
         If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
            TempStr = FormatNumberToNull(Amt)
         ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Then
            Amt = m_BillingDoc.PAID_AMOUNT
            TempStr = FormatNumberToNull(Amt)
         Else
            Amt = 0
            TempStr = ""
         End If
         Total1(5) = Total1(5) + Amt
         Total2(5) = Total2(5) + Amt
         SumRow = SumRow - Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(6)
         If m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
            Amt = m_BillingDoc.PAID_AMOUNT
            TempStr = FormatNumberToNull(Amt)
         Else
            Amt = 0
            TempStr = ""
         End If
         Total1(6) = Total1(6) + Amt
         Total2(6) = Total2(6) + Amt
         SumRow = SumRow + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(7)
         If m_BillingDoc.DOCUMENT_TYPE = RECEIPT2_DOCTYPE Then
            Amt = m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT - m_BillingDoc.CREDIT_AMOUNT
            TempStr = FormatNumberToNull(Amt)
         Else
            Amt = 0
            TempStr = ""
         End If
         Total1(7) = Total1(7) + Amt
         Total2(7) = Total2(7) + Amt
         SumRow = SumRow - Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(8)
         Amt = SumRow
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(9)
         Amt = m_BillingDoc.CASH_PMT
         TempStr = FormatNumberToNull(Amt, , , "")
         Total1(9) = Total1(9) + Amt
         Total2(9) = Total2(9) + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(10)
         Amt = m_BillingDoc.BANKTRF_PMT
         TempStr = FormatNumberToNull(Amt, , , "")
         Total1(10) = Total1(10) + Amt
         Total2(10) = Total2(10) + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      
         Set HeadCf = m_Headers1.Fields(11)
         Amt = m_BillingDoc.CHEQUE_PMT
         TempStr = FormatNumberToNull(Amt, , , "")
         Total1(11) = Total1(11) + Amt
         Total2(11) = Total2(11) + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(12)
         'TempStr = m_BillingDoc.CHEQUE_NO")
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(13)
         'TempStr = DateToStringExtEx2(m_BillingDoc.CHEQUE_DATE"))
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(14)
         'TempStr = DateToStringExtEx2(m_BillingDoc.EFFECTIVE_DATE"))
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(15)
         'TempStr = m_BillingDoc.DOC_ID_NO")
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(16)
         'TempStr = DateToStringExtEx2(m_BillingDoc.DOC_ID_DATE"))
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(17)
         'TempStr = FormatNumberToNull(m_BillingDoc.DOC_ID_PAID"), , , "")
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      
      End If
      
      If Len(m_BillingDoc.CHEQUE_NO) > 0 Then
         If TempCheque <> m_BillingDoc.CHEQUE_NO Then
            If Len(TempChequeNo) <= 0 Then
               TempChequeNo = m_BillingDoc.CHEQUE_NO
               TempChequeDate = DateToStringExtEx2(m_BillingDoc.CHEQUE_DATE)
               TempChequeEffectdate = DateToStringExtEx2(m_BillingDoc.EFFECTIVE_DATE)
            Else
               TempChequeNo = TempChequeNo & vbCrLf & m_BillingDoc.CHEQUE_NO
               TempChequeDate = TempChequeDate & vbCrLf & DateToStringExtEx2(m_BillingDoc.CHEQUE_DATE)
               TempChequeEffectdate = TempChequeEffectdate & vbCrLf & DateToStringExtEx2(m_BillingDoc.EFFECTIVE_DATE)
            End If
            TempCheque = m_BillingDoc.CHEQUE_NO
         End If
      End If
      
      If Len(m_BillingDoc.DOC_ID_NO) > 0 Then
         Set BD = GetObject("CBillingDoc", TempColl, m_BillingDoc.DOC_ID_NO, False)
         If BD Is Nothing Then
            If Len(TempDocNo) <= 0 Then
               Set TempColl = New Collection
               TempDocNo = m_BillingDoc.DOC_ID_NO
               TempDocDate = DateToStringExtEx2(m_BillingDoc.DOC_ID_DATE)
               If m_BillingDoc.DOCUMENT_TYPE = RECEIPT2_DOCTYPE Then
                  If m_BillingDoc.DOCUMENT_TYPE_RCP = INVOICE_DOCTYPE Then
                     TempDocPaid = FormatNumber(m_BillingDoc.DOC_ID_PAID)
                  ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = CN_DOCTYPE Then
                     TempDocPaid = FormatNumber(-1 * m_BillingDoc.CNDN_AMOUNT)         'ลดหนี้
                  ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = DN_DOCTYPE Then
                     TempDocPaid = FormatNumber(m_BillingDoc.CNDN_AMOUNT)         'เพิ่มหนี้
                  ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = RETURN_DOCTYPE Then
                     TempDocPaid = FormatNumber(-1 * m_BillingDoc.DOC_ID_PAID)
                  End If
               End If
            Else
               TempDocNo = TempDocNo & vbCrLf & m_BillingDoc.DOC_ID_NO
               TempDocDate = TempDocDate & vbCrLf & DateToStringExtEx2(m_BillingDoc.DOC_ID_DATE)
               If m_BillingDoc.DOCUMENT_TYPE = RECEIPT2_DOCTYPE Then
                  If m_BillingDoc.DOCUMENT_TYPE_RCP = INVOICE_DOCTYPE Then
                     TempDocPaid = TempDocPaid & vbCrLf & FormatNumber(m_BillingDoc.DOC_ID_PAID)
                  ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = CN_DOCTYPE Then
                     TempDocPaid = TempDocPaid & vbCrLf & FormatNumber(-1 * m_BillingDoc.CNDN_AMOUNT)     'ลดหนี้
                  ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = DN_DOCTYPE Then
                     TempDocPaid = TempDocPaid & vbCrLf & FormatNumber(m_BillingDoc.CNDN_AMOUNT)     'เพิ่มหนี้
                  ElseIf m_BillingDoc.DOCUMENT_TYPE_RCP = RETURN_DOCTYPE Then
                     TempDocPaid = TempDocPaid & vbCrLf & FormatNumber(-1 * m_BillingDoc.DOC_ID_PAID)
                  End If
               End If
            End If
            Set BD = New CBillingDoc
            BD.DOC_ID_NO = m_BillingDoc.DOC_ID_NO
            Call TempColl.add(BD, m_BillingDoc.DOC_ID_NO)
            Set BD = Nothing
         End If
      End If
      
      PrevKey3 = m_BillingDoc.DOCUMENT_NO
      
      Rs.MoveNext
   Wend
   
   If I <> 1 And Len(TempChequeNo) > 0 Then
      Call GenerateCheque(TempChequeNo, TempChequeDate, TempChequeEffectdate, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
      TempChequeNo = ""
      TempChequeDate = ""
      TempChequeEffectdate = ""
   End If
   
   If I <> 1 And Len(TempDocNo) > 0 Then
      Call GenerateDoc(TempDocNo, TempDocDate, TempDocPaid, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      TempDocNo = ""
      TempDocDate = ""
      TempDocPaid = ""
      Set TempColl = Nothing
   End If
   
   Call GenerateFooter("", "รวม " & PrevKey2, Total1, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone
   
   Call GenerateFooter("", "สุทธิ", Total2, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   Vsp.TableBorder = tbNone

   
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

Private Sub GenerateFooter(txt1 As String, txt2 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
   Call BodyCf.SetFieldValue(TempWidth, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j >= 4 Then
         TempStr = FormatNumberToNull(Val(Tot(j)), , , "")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
   Next HeadCf
   
   Set BodyCf = Nothing
End Sub
Private Sub GenerateBalance(txt1 As String, Balance As Double, Details As CFieldList)
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
      ElseIf j = 8 Then
         TempStr = FormatNumberToNull(Balance)
      Else
         TempStr = ""
      End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
   Next HeadCf
   
   Set BodyCf = Nothing
End Sub
Private Sub GenerateCheque(txt1 As String, txt2 As String, Txt3 As String, Details As CFieldList)
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
      If j = 12 Then
         TempStr = txt1
      ElseIf j = 13 Then
         TempStr = txt2
      ElseIf j = 14 Then
         TempStr = Txt3
      Else
         TempStr = ""
      End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
   Next HeadCf
   
   Set BodyCf = Nothing
End Sub
Private Sub GenerateDoc(txt1 As String, txt2 As String, Txt3 As String, Details As CFieldList)
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
      If j = 15 Then
         TempStr = txt1
      ElseIf j = 16 Then
         TempStr = txt2
      ElseIf j = 17 Then
         TempStr = Txt3
      Else
         TempStr = ""
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
