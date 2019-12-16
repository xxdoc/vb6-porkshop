VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling018"
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
Dim FromDate As Date
Dim ToDate As Date
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   If mcolParam("SHOW_SUMMARY") = 1 Then
      Call Cf.SetFieldValue(8, "^", "รายละเอียด", ">")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(10, "^", "รายละเอียดสินค้า", ">")
      Call m_Headers0.AddField(Cf)
   End If
   
   If mcolParam("SHOW_SUMMARY") = 1 Then
      Call Cf.SetFieldValue(2, "^", "รหัส", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "ลูกค้า", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "", "<")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(2, "^", "รหัสสินค้า", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(6, "^", "รายละเอียด", "<")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "หน่วย", "^")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("NOT_SHOW_AMOUNT") = 0 Then
      FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
      ToDate = DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1)
      
      Call Cf.SetFieldValue(2 * (DateDiff("M", FromDate, ToDate) + 2), "^", "จำนวน", ">")
      Call m_Headers0.AddField(Cf)
      
      '------------------------------------------------------------------------------------------>          ยอดขาย
      While FromDate <= ToDate
         Call Cf.SetFieldValue(2, "^", Month(FromDate) & "/" & Year(FromDate) + 543, ">")
         Call m_Headers1.AddField(Cf)
         
         FromDate = DateAdd("M", 1, FromDate)
      Wend
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers1.AddField(Cf)
   End If
   '------------------------------------------------------------------------------------------>
   
   If mcolParam("NOT_SHOW_PRICE") = 0 Then
      FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
      ToDate = DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1)
      
      Call Cf.SetFieldValue(2 * (DateDiff("M", FromDate, ToDate) + 2), "^", "มูลค่า", ">")
      Call m_Headers0.AddField(Cf)
      '------------------------------------------------------------------------------------------>          ยอดขาย
      While FromDate <= ToDate
         Call Cf.SetFieldValue(2, "^", Month(FromDate) & "/" & Year(FromDate) + 543, ">")
         Call m_Headers1.AddField(Cf)
            
         FromDate = DateAdd("M", 1, FromDate)
      Wend
      
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers1.AddField(Cf)
   End If
   '------------------------------------------------------------------------------------------>
   
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
   Vsp.MarginFooter = 200
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 500
   Vsp.MarginRight = 500
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1), TempFromDate, TempToDate)
   TempFromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   
   Call GetSaleAmountMonthByCustomerStockCode(TotalMoves, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   FromDate = "จากงวด " & mcolParam("FROM_MONTH_ID") & "/" & mcolParam("FROM_YEAR_NO")
   ToDate = "ถึง " & mcolParam("TO_MONTH_ID") & "/" & mcolParam("TO_YEAR_NO")
   
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(30) As Double
Dim Total2(30) As Double
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim TotalMove As CBillingDoc
Dim m_BillingDoc As CBillingDoc

Dim FromDate As Date
Dim ToDate As Date
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
   
   Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1), FromDate, ToDate)
   FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   
   m_BillingDoc.FROM_DATE = FromDate
   m_BillingDoc.TO_DATE = ToDate
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   
   m_BillingDoc.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
   m_BillingDoc.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
   Call m_BillingDoc.QueryData(1008, Rs, iCount)
   
   I = 0
   
   While Not Rs.EOF
      I = I + 1
      Call m_Details1.ClearField
      Call m_BillingDoc.PopulateFromRS(1008, Rs)
      
      
      If PrevKey1 <> m_BillingDoc.APAR_CODE And I <> 1 Then
         If mcolParam("SHOW_SUMMARY") = 0 Then
            Call GenerateFooter("รวม ", PrevKey2, Total1, m_Details1, 0)
         Else
            Call GenerateFooter(PrevKey1, PrevKey2, Total1, m_Details1, 0)
         End If
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         If mcolParam("SHOW_SUMMARY") = 0 Then
            Vsp.FontBold = True
         End If
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         If mcolParam("SHOW_SUMMARY") = 0 Then
            Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
         End If
         ' รวม
         For j = 1 To UBound(Total1)
            Total1(j) = 0
         Next j
      ElseIf I = 1 Then
         If mcolParam("SHOW_SUMMARY") = 0 Then
            Vsp = "ลูกค้า  :  " & m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
         End If
      End If

      PrevKey1 = m_BillingDoc.APAR_CODE
      PrevKey2 = m_BillingDoc.APAR_NAME
      SumRow = 0
      Call m_Details1.ClearField
      
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = m_BillingDoc.STOCK_NO
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = m_BillingDoc.STOCK_DESC
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(3)
      TempStr = m_BillingDoc.UNIT_CHANGE_NAME
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = 3
      If mcolParam("NOT_SHOW_AMOUNT") = 0 Then
            FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
            ToDate = DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1)
               
            While FromDate <= ToDate
               Set TotalMove = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO & "-" & Year(FromDate) & "-" & Format(Month(FromDate), "00")) & "-" & INVOICE_DOCTYPE, True)
               Amt = TotalMove.TOTAL_AMOUNT
               Set TotalMove = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO & "-" & Year(FromDate) & "-" & Format(Month(FromDate), "00")) & "-" & RECEIPT1_DOCTYPE, True)
               Amt = Amt + TotalMove.TOTAL_AMOUNT
               Set TotalMove = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO & "-" & Year(FromDate) & "-" & Format(Month(FromDate), "00")) & "-" & RETURN_DOCTYPE, True)
               Amt = Amt - TotalMove.TOTAL_AMOUNT
               
               j = j + 1
               Set HeadCf = m_Headers1.Fields(j)
               TempStr = FormatNumberToNull(Amt)
               Total1(j) = Total1(j) + Amt
               Total2(j) = Total2(j) + Amt
               SumRow = SumRow + Amt
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               FromDate = DateAdd("M", 1, FromDate)
            Wend
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(SumRow)
            Total1(j) = Total1(j) + SumRow
            Total2(j) = Total2(j) + SumRow
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
      End If
      SumRow = 0
      If mcolParam("NOT_SHOW_PRICE") = 0 Then
         FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
         ToDate = DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1)
   
         While FromDate <= ToDate
            Set TotalMove = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO & "-" & Year(FromDate) & "-" & Format(Month(FromDate), "00")) & "-" & INVOICE_DOCTYPE, True)
            If mcolParam("NOT_DISCOUNT") = "0" Then
               Amt = TotalMove.TOTAL_PRICE - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT
            Else
               Amt = TotalMove.TOTAL_PRICE
            End If
            Set TotalMove = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO & "-" & Year(FromDate) & "-" & Format(Month(FromDate), "00")) & "-" & RECEIPT1_DOCTYPE, True)
            If mcolParam("NOT_DISCOUNT") = "0" Then
               Amt = Amt + TotalMove.TOTAL_PRICE - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT
            Else
               Amt = Amt + TotalMove.TOTAL_PRICE
            End If
            Set TotalMove = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO & "-" & Year(FromDate) & "-" & Format(Month(FromDate), "00")) & "-" & RETURN_DOCTYPE, True)
            If mcolParam("NOT_DISCOUNT") = "0" Then
               Amt = Amt - (TotalMove.TOTAL_PRICE - TotalMove.DISCOUNT_AMOUNT - TotalMove.EXT_DISCOUNT_AMOUNT)
            Else
               Amt = Amt - TotalMove.TOTAL_PRICE
            End If
            
            j = j + 1
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(Amt)
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            FromDate = DateAdd("M", 1, FromDate)
         Wend
         
         j = j + 1
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(SumRow)
         Total1(j) = Total1(j) + SumRow
         Total2(j) = Total2(j) + SumRow
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      End If
      
      If mcolParam("SHOW_SUMMARY") = 0 Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
      Rs.MoveNext
   Wend
   
   If mcolParam("SHOW_SUMMARY") = 0 Then
      Call GenerateFooter("รวม", PrevKey2, Total1, m_Details1, 0)
   Else
      Call GenerateFooter(PrevKey1, PrevKey2, Total1, m_Details1, 0)
   End If
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   If mcolParam("SHOW_SUMMARY") = 0 Then
      Vsp.FontBold = True
   End If
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
         
   Call GenerateFooter("สุทธิ", "", Total2, m_Details1, 0)
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
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.Description
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(txt1 As String, txt2 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j > 2 Then
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
Dim dy(0 To 1) As Double
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
   dy(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage & " / " & glbUser.USER_NAME)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.TAX_ID)
   Vsp.CurrentY = dy(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   If mblnHeader Then
      If Vsp.CurrentPage > 1 Then
         Call printHeader
      End If
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
   If mcolParam("SHOW_SUMMARY") > 0 Then
      Vsp.Paragraph = "*สรุป"
   End If
End Sub

