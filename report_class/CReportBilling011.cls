VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling0011"
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
   
   If mcolParam("SUMMARY") = 0 Then
      Call Cf.SetFieldValue(2, "^", "วันที่", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(3, "^", "เลขที่", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "", "^")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(5, "^", "ลูกค้า", "<")
      Call m_Headers1.AddField(Cf)

   End If
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ราคา/หน่วย", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ส่วนลด", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รวมเงิน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ส่วนลดรวม", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดขายสุทธิ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "อ้างอิง", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "หมายเหตุ", "<")
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
    'g_CharSpacing = -0 / Vsp.TwipsPerPixelX
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Vsp.ZoomMode = zmPageWidth
   
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & "        " & FromDate & " " & ToDate & vbCrLf & _
                                       "จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " จากคลัง " & EmptyToString(mcolParam("FROM_LOCATION_NO"), "N/A") & " ถึงคลัง " & EmptyToString(mcolParam("TO_LOCATION_NO"), "N/A") & " จาก พข " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึง พข " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
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
Dim Total1(10) As Double
Dim Total2(10) As Double
Dim Total3(10) As Double
Dim Total4(10) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim k As Byte
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
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField
      
      Set m_BillingDoc = New CBillingDoc
      
      m_BillingDoc.BILLING_DOC_ID = -1
      If mcolParam("CONSIGNMENT") = 1 Then
         m_BillingDoc.DOCUMENT_TYPE_SET = "(" & PO_DOCTYPE & ")"
         m_BillingDoc.FROM_DUE_DATE = mcolParam("FROM_BILL_DATE")
         m_BillingDoc.TO_DUE_DATE = mcolParam("TO_BILL_DATE")
         m_BillingDoc.CONSIGNMENT_FLAG = "Y"
      Else
         m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
         m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
         m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
      End If
      m_BillingDoc.APAR_IND = 1
      m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
      m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
      m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
      m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
      m_BillingDoc.FROM_LOCATION_NO = mcolParam("FROM_LOCATION_NO")
      m_BillingDoc.TO_LOCATION_NO = mcolParam("TO_LOCATION_NO")
      m_BillingDoc.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      m_BillingDoc.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      m_BillingDoc.FREE_FLAG = StringToFreeFlag(mcolParam("INCLUDE_FREE"))
      Call m_BillingDoc.QueryData(19, Rs, iCount)
      
      I = 0
      
      While Not Rs.EOF
         I = I + 1
         Call m_BillingDoc.PopulateFromRS(19, Rs)
         
         Amt = 0
         If PrevKey6 <> m_BillingDoc.APAR_CODE And I > 1 Then
            If mcolParam("SUMMARY") = 0 Then
               Call GenerateFooter("                        รวม", PrevKey4, "", Total1, m_Details1, 0, PrevKey5)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               Call GenerateFooter("            รวม", PrevKey2, "", Total2, m_Details1, 0, PrevKey5)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               Call GenerateFooter("รวม", PrevKey7, "", Total3, m_Details1, 0, PrevKey5)
               Vsp.FontBold = True
            Else
               Call GenerateFooter("", PrevKey7 & "/" & PrevKey6, "", Total3, m_Details1, 0, PrevKey5)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
               Total3(j) = 0
            Next j
            If mcolParam("SUMMARY") = 0 Then
               Vsp = m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.APAR_CODE & " )"
               Vsp = "           " & m_BillingDoc.STOCK_DESC & " ( " & m_BillingDoc.STOCK_NO & " )"
            End If
         ElseIf PrevKey1 <> m_BillingDoc.STOCK_NO And I > 1 Then
            If mcolParam("SUMMARY") = 0 Then
               Call GenerateFooter("                        รวม", PrevKey4, "", Total1, m_Details1, 0, PrevKey5)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               Call GenerateFooter("            รวม", PrevKey2, "", Total2, m_Details1, 0, PrevKey5)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
            Next j
            If mcolParam("SUMMARY") = 0 Then
               Vsp = "           " & m_BillingDoc.STOCK_DESC & " ( " & m_BillingDoc.STOCK_NO & " )"
            End If
         ElseIf PrevKey3 <> Trim(Str(m_BillingDoc.DOCUMENT_TYPE)) And I <> 1 Then
            If mcolParam("SUMMARY") = 0 Then
               Call GenerateFooter("                        รวม", PrevKey4, "", Total1, m_Details1, 0, PrevKey5)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            
         ElseIf I = 1 Then
            If mcolParam("SUMMARY") = 0 Then
               Vsp = m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.APAR_CODE & " )"
               Vsp = "           " & m_BillingDoc.STOCK_DESC & " ( " & m_BillingDoc.STOCK_NO & " )"
            End If
         End If
         
         PrevKey1 = m_BillingDoc.STOCK_NO
         PrevKey2 = m_BillingDoc.STOCK_DESC
         PrevKey3 = Trim(Str(m_BillingDoc.DOCUMENT_TYPE))
         PrevKey4 = SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE)
         PrevKey5 = m_BillingDoc.UNIT_CHANGE_NAME
         PrevKey6 = m_BillingDoc.APAR_CODE
         PrevKey7 = m_BillingDoc.APAR_NAME
         
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = m_BillingDoc.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = 2
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         Amt = MyDiffEx(m_BillingDoc.ITEM_AMOUNT, m_BillingDoc.UNIT_MULTIPLE)
         If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = -Amt
         End If
         TempStr = FormatNumberToNull(Amt) & " " & m_BillingDoc.UNIT_TRAN_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Total1(k) = Total1(k) - m_BillingDoc.ITEM_AMOUNT
            Total2(k) = Total2(k) - m_BillingDoc.ITEM_AMOUNT
            Total3(k) = Total3(k) - m_BillingDoc.ITEM_AMOUNT
            Total4(k) = Total4(k) - m_BillingDoc.ITEM_AMOUNT
         Else
            Total1(k) = Total1(k) + m_BillingDoc.ITEM_AMOUNT
            Total2(k) = Total2(k) + m_BillingDoc.ITEM_AMOUNT
            Total3(k) = Total3(k) + m_BillingDoc.ITEM_AMOUNT
            Total4(k) = Total4(k) + m_BillingDoc.ITEM_AMOUNT
         End If
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(m_BillingDoc.AVG_PRICE * m_BillingDoc.UNIT_MULTIPLE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = -m_BillingDoc.TOTAL_PRICE
         Else
            Amt = m_BillingDoc.TOTAL_PRICE
         End If
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(m_BillingDoc.DISCOUNT_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + m_BillingDoc.DISCOUNT_AMOUNT
         Total2(k) = Total2(k) + m_BillingDoc.DISCOUNT_AMOUNT
         Total3(k) = Total3(k) + m_BillingDoc.DISCOUNT_AMOUNT
         Total4(k) = Total4(k) + m_BillingDoc.DISCOUNT_AMOUNT
         
         If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = Amt + m_BillingDoc.DISCOUNT_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = Amt - m_BillingDoc.DISCOUNT_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
            Amt = Amt - m_BillingDoc.DISCOUNT_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = PO_DOCTYPE Then
            Amt = Amt - m_BillingDoc.DISCOUNT_AMOUNT
         End If
         
         k = k + 1
          Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
          Total3(k) = Total3(k) + Amt
          Total4(k) = Total4(k) + Amt
          
          k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = EmptyToString(FormatNumber(m_BillingDoc.EXT_DISCOUNT_PERCENT) & "%", "")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         If m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            Amt = Amt + m_BillingDoc.EXT_DISCOUNT_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = Amt - m_BillingDoc.EXT_DISCOUNT_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
            Amt = Amt - m_BillingDoc.EXT_DISCOUNT_AMOUNT
         ElseIf m_BillingDoc.DOCUMENT_TYPE = PO_DOCTYPE Then
            Amt = Amt - m_BillingDoc.EXT_DISCOUNT_AMOUNT
         End If
         TempStr = FormatNumberToNull(Amt)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
          Total3(k) = Total3(k) + Amt
          Total4(k) = Total4(k) + Amt
          
          k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         If Len(m_BillingDoc.REFER_TEXT) > 0 Then
            TempStr = m_BillingDoc.REFER_TEXT
         Else
            TempStr = m_BillingDoc.PO_NO
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = m_BillingDoc.NOTE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If mcolParam("SUMMARY") = 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         Rs.MoveNext
      Wend
            
      If mcolParam("SUMMARY") = 0 Then
         Call GenerateFooter("                        รวม", PrevKey4, "", Total1, m_Details1, 0, PrevKey5)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Call GenerateFooter("            รวม", PrevKey2, "", Total2, m_Details1, 0, PrevKey5)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Call GenerateFooter("รวม", PrevKey7, "", Total3, m_Details1, 0, PrevKey5)
         Vsp.FontBold = True
      Else
         Call GenerateFooter("", PrevKey7 & "/" & PrevKey6, "", Total3, m_Details1, 0, PrevKey5)
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Vsp = ""
      
      Call GenerateFooter("สุทธิ", "", "", Total4, m_Details1, 0)
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

Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long, Optional UnitChangeName As String)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   For j = 1 To 2
      Set HeadCf = m_Headers1.Fields(j)
      TempWidth = TempWidth + HeadCf.Width
   Next j
   TempStr = "" & txt1 & " " & txt2 & " " & Txt3
   Call BodyCf.SetFieldValue(TempWidth, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 3 Or j = 6 Or j = 8 Then
         TempStr = FormatNumberToNull(Tot(j))
         If j = 3 Then
            TempStr = TempStr & " " & UnitChangeName
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      ElseIf j < 3 Then
      Else
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
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub

