VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling030"
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

Private TotalMoves As Collection
Private TotalMoveCmps As Collection
Private TotalMoveRts As Collection
Private TotalMoveCmpRts As Collection

Private DistinctStockNos As Collection
Private OtherMoves As Collection
Private OtherMoveCmps As Collection
Private DistinctStockNosConsignment As Collection
Private CONSIGNMENT As Collection
Private ConsignmentCmps As Collection
Private NotStockNo As Collection

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
   
   Set TotalMoves = New Collection
   Set TotalMoveCmps = New Collection
   Set TotalMoveRts = New Collection
   Set TotalMoveCmpRts = New Collection
   
   Set OtherMoves = New Collection
   Set OtherMoveCmps = New Collection
   Set CONSIGNMENT = New Collection
   Set ConsignmentCmps = New Collection
   Set DistinctStockNos = New Collection
   Set DistinctStockNosConsignment = New Collection
   Set NotStockNo = New Collection
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
   
   Set TotalMoves = Nothing
   Set TotalMoveCmps = Nothing
   Set TotalMoveRts = Nothing
   Set TotalMoveCmpRts = Nothing
   Set OtherMoves = Nothing
   Set OtherMoveCmps = Nothing
   Set CONSIGNMENT = Nothing
   Set ConsignmentCmps = Nothing
   Set DistinctStockNos = Nothing
   Set DistinctStockNosConsignment = Nothing
   Set NotStockNo = Nothing
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
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(6, "^", "รายการเอียด", "<")
   Call m_Headers0.AddField(Cf)
   If Month(mcolParam("FROM_RCP_DATE")) = Month(mcolParam("TO_RCP_DATE")) Then
      Call Cf.SetFieldValue(6, "^", "ยอดขายวันที่ " & Format(Day(mcolParam("FROM_RCP_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_RCP_DATE")), ">")
   ElseIf Year(mcolParam("FROM_RCP_DATE")) = Year(mcolParam("TO_RCP_DATE")) Then
      Call Cf.SetFieldValue(6, "^", "ยอดขายวันที่ " & Format(Day(mcolParam("FROM_RCP_DATE")), "00") & "/" & Format(Month(mcolParam("TO_RCP_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   Else
      Call Cf.SetFieldValue(6, "^", "ยอดขายวันที่ " & DateToStringExtEx2(mcolParam("FROM_RCP_DATE")) & "-" & DateToStringExtEx2(mcolParam("TO_RCP_DATE")), ">")
   End If
   Call m_Headers0.AddField(Cf)
   If Month(mcolParam("FROM_BILL_DATE")) = Month(mcolParam("TO_BILL_DATE")) Then
      Call Cf.SetFieldValue(6, "^", "ยอดขายสะสม " & Format(Day(mcolParam("FROM_BILL_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   ElseIf Year(mcolParam("FROM_BILL_DATE")) = Year(mcolParam("TO_BILL_DATE")) Then
      Call Cf.SetFieldValue(6, "^", "ยอดขายสะสม " & Format(Day(mcolParam("FROM_BILL_DATE")), "00") & "/" & Format(Month(mcolParam("FROM_BILL_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   Else
      Call Cf.SetFieldValue(6, "^", "ยอดขายสะสม " & DateToStringExtEx2(mcolParam("FROM_BILL_DATE")) & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   End If
   Call m_Headers0.AddField(Cf)
   If mcolParam("SHOW_COST") > 0 Then
      If Month(mcolParam("FROM_BILL_DATE")) = Month(mcolParam("TO_BILL_DATE")) Then
         Call Cf.SetFieldValue(4, "^", "ต้นทุน " & Format(Day(mcolParam("FROM_BILL_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
      ElseIf Year(mcolParam("FROM_BILL_DATE")) = Year(mcolParam("TO_BILL_DATE")) Then
         Call Cf.SetFieldValue(4, "^", "ต้นทุน " & Format(Day(mcolParam("FROM_BILL_DATE")), "00") & "/" & Format(Month(mcolParam("FROM_BILL_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
      Else
         Call Cf.SetFieldValue(4, "^", "ต้นทุน " & DateToStringExtEx2(mcolParam("FROM_BILL_DATE")) & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
      End If
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "ต้นทุน", ">")
      Call m_Headers0.AddField(Cf)
   End If
   Call Cf.SetFieldValue(4, "^", "กำไร", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "รายการสินค้า/รหัส", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_COST") > 0 Then
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">", , True)
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "จำนวนเงิน", ">", , True)
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(0, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
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
Dim FromDate As Date
Dim ToDate As Date
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call LoadStockCodeFromTo(Nothing, NotStockNo, mcolParam("NOT_FROM_STOCK_NO"), mcolParam("NOT_TO_STOCK_NO"), 1)
   
   Call GetSaleAmountAparGroupDocTypeStockCodeFree(TotalMoves, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountAparGroupDocTypeStockCodeFree(TotalMoveCmps, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   
   If mcolParam("FROM_BILL_DATE") >= mcolParam("FROM_RCP_DATE") Then
      FromDate = mcolParam("FROM_RCP_DATE")
   ElseIf mcolParam("FROM_BILL_DATE") <= mcolParam("FROM_RCP_DATE") Then
      FromDate = mcolParam("FROM_BILL_DATE")
   End If
   If mcolParam("TO_BILL_DATE") >= mcolParam("TO_RCP_DATE") Then
      ToDate = mcolParam("TO_BILL_DATE")
   ElseIf mcolParam("TO_BILL_DATE") <= mcolParam("TO_RCP_DATE") Then
      ToDate = mcolParam("TO_RCP_DATE")
   End If
   
   If (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
      Call GetDistinctStockCode(DistinctStockNos, FromDate, ToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
      Call GetSaleAmountStockCodeDocType(OtherMoves, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
      Call GetSaleAmountStockCodeDocType(OtherMoveCmps, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   End If
   If mcolParam("CONSIGNMENT") = 1 Then
      Call GetDistinctStockCode(DistinctStockNosConsignment, FromDate, ToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & PO_DOCTYPE & ")", mcolParam("INCLUDE_FREE"), 1)
      Call GetSaleAmountStockCodeDocType(CONSIGNMENT, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & PO_DOCTYPE & ")", mcolParam("INCLUDE_FREE"), 1)
      Call GetSaleAmountStockCodeDocType(ConsignmentCmps, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & PO_DOCTYPE & ")", mcolParam("INCLUDE_FREE"), 1)
   End If
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       "จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " ไม่รวมสินค้าจาก " & EmptyToString(mcolParam("NOT_FROM_STOCK_NO"), "N/A") & " ถึง " & EmptyToString(mcolParam("NOT_TO_STOCK_NO"), "N/A")
                                       
                                       
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
Dim TempBillingDoc As CBillingDoc
Dim TempBillingDoc2 As CBillingDoc
Dim TempBillingDoc3 As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(20) As Double
Dim Total2(20) As Double
Dim Total3(20) As Double
Dim Total4(20) As Double
Dim SumItem As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim TotalPrice As Double
Dim CostAmount As Double
Dim CheckStock As Boolean
Dim FromDate As Date
Dim ToDate As Date

Dim SaleReturnColl As Collection
Dim TempSaleReturn As CBillingDoc
Dim k As Long
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set SaleReturnColl = New Collection
      
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
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
   
   MasterInd = "76"
   
   If mcolParam("PRINT_TO_EXCEL") = 1 Then
      Call printHeader
   End If
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   If mcolParam("FROM_BILL_DATE") >= mcolParam("FROM_RCP_DATE") Then
      FromDate = mcolParam("FROM_RCP_DATE")
   ElseIf mcolParam("FROM_BILL_DATE") <= mcolParam("FROM_RCP_DATE") Then
      FromDate = mcolParam("FROM_BILL_DATE")
   End If
   If mcolParam("TO_BILL_DATE") >= mcolParam("TO_RCP_DATE") Then
      ToDate = mcolParam("TO_BILL_DATE")
   ElseIf mcolParam("TO_BILL_DATE") <= mcolParam("TO_RCP_DATE") Then
      ToDate = mcolParam("TO_RCP_DATE")
   End If
   m_BillingDoc.FROM_DATE = FromDate
   m_BillingDoc.TO_DATE = ToDate
   m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.FREE_FLAG = StringToFreeFlag(mcolParam("INCLUDE_FREE"))
   Call m_BillingDoc.QueryData(76, Rs, iCount)
   
   I = 0
   
   While Not Rs.EOF
      MasterInd = "76"
      Set m_BillingDoc = New CBillingDoc
      
      Call m_BillingDoc.PopulateFromRS(76, Rs)
      
      Call CheckNotStock(m_BillingDoc.STOCK_NO, CheckStock)
      If CheckStock = True Then
         If m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then       ' Merge Record ของใบขายสดเข้ากับใบส่งของ
            Set TempSaleReturn = GetObject("CBillingDoc", SaleReturnColl, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), False)
            If TempSaleReturn Is Nothing Then
               m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE
               Call SaleReturnColl.add(m_BillingDoc, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO))
            End If
         Else
            Set TempSaleReturn = GetObject("CBillingDoc", SaleReturnColl, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & m_BillingDoc.DOCUMENT_TYPE & "-" & m_BillingDoc.STOCK_NO), False)
            If TempSaleReturn Is Nothing Then
               Call SaleReturnColl.add(m_BillingDoc, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & m_BillingDoc.DOCUMENT_TYPE & "-" & m_BillingDoc.STOCK_NO))
            End If
         End If
         Set m_BillingDoc = Nothing
      End If
      Rs.MoveNext
   Wend
   
   For Each m_BillingDoc In SaleReturnColl
      I = I + 1
      Call m_Details1.ClearField
      If ((mcolParam("STOCK_TYPE") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE1") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE2") <> m_BillingDoc.STOCK_TYPE)) And (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
         I = I - 1
      Else
         If PrevKey1 <> m_BillingDoc.APAR_GROUP_NAME And I <> 1 Then
            Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             If mcolParam("SHOW_DETAIL") = 1 Then
               Vsp.FontBold = True
               Vsp.TextColor = RGB(0, 0, 255)
             End If
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
            Vsp.TextColor = RGB(0, 0, 0)
            
            Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Vsp.TextColor = RGB(0, 0, 255)
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
             Vsp.TextColor = RGB(0, 0, 0)
            
             Vsp = "กลุ่มลูกค้า  :  " & m_BillingDoc.APAR_GROUP_NAME
             If mcolParam("SHOW_DETAIL") = 1 And (m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or mcolParam("INCLUDE_RETURN") = 1) Then
               Vsp = "       ประเภท  :  " & SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE)
             End If
             ' รวม
             For j = 1 To UBound(Total1)
                Total1(j) = 0
                Total2(j) = 0
             Next j
             
          ElseIf PrevKey2 <> Trim(Str(m_BillingDoc.DOCUMENT_TYPE)) And I <> 1 Then
             Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             If mcolParam("SHOW_DETAIL") = 1 Then
               Vsp.FontBold = True
               Vsp.TextColor = RGB(0, 0, 255)
             End If
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
            Vsp.TextColor = RGB(0, 0, 0)
            
            If mcolParam("SHOW_DETAIL") = 1 And (m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or mcolParam("INCLUDE_RETURN") = 1) Then
               Vsp = "       ประเภท  :  " & SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE)
            End If
             ' รวม
             For j = 1 To UBound(Total1)
                Total1(j) = 0
             Next j
          ElseIf I = 1 Then
            Vsp.TextColor = RGB(0, 0, 0)
             Vsp = "กลุ่มลูกค้า  :  " & m_BillingDoc.APAR_GROUP_NAME
             If mcolParam("SHOW_DETAIL") = 1 And (m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or mcolParam("INCLUDE_RETURN") = 1) Then
               Vsp = "       ประเภท  :  " & SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE)
            End If
          End If
          
          PrevKey1 = m_BillingDoc.APAR_GROUP_NAME
          PrevKey2 = Trim(Str(m_BillingDoc.DOCUMENT_TYPE))
          
          Call m_Details1.ClearField
         Vsp.TextColor = RGB(0, 0, 0)
         
         If m_BillingDoc.STOCK_NO = "44-2001" Then
            Debug.Print
         End If
         
          Set HeadCf = m_Headers1.Fields(1)
          TempStr = m_BillingDoc.STOCK_DESC & " /" & m_BillingDoc.STOCK_NO
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Set TempBillingDoc = GetObject("CBillingDoc", TotalMoveCmps, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
            Set TempSaleReturn = GetObject("CBillingDoc", TotalMoveCmps, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & RECEIPT1_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
          ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then ' ใบรับคืน
            Set TempBillingDoc = GetObject("CBillingDoc", TotalMoveCmps, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & RETURN_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
          End If
          Set HeadCf = m_Headers1.Fields(2)
          If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            SumItem = TempBillingDoc.TOTAL_AMOUNT + TempSaleReturn.TOTAL_AMOUNT
          ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            SumItem = -TempBillingDoc.TOTAL_AMOUNT
          End If
          TempStr = FormatNumberToNull(SumItem)
          Total1(2) = Total1(2) + SumItem
          Total2(2) = Total2(2) + SumItem
          Total3(2) = Total3(2) + SumItem
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Amt = Amt + TempSaleReturn.TOTAL_PRICE - TempSaleReturn.DISCOUNT_AMOUNT - TempSaleReturn.EXT_DISCOUNT_AMOUNT
         Else
            Amt = -(TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
         End If
          Set HeadCf = m_Headers1.Fields(3)
          TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
          Total1(3) = Total1(3) + MyDiffEx(Amt, SumItem)
          Total2(3) = Total2(3) + MyDiffEx(Amt, SumItem)
          Total3(3) = Total3(3) + MyDiffEx(Amt, SumItem)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set HeadCf = m_Headers1.Fields(4)
          TempStr = FormatNumberToNull(Amt)
          Total1(4) = Total1(4) + Amt
          Total2(4) = Total2(4) + Amt
          Total3(4) = Total3(4) + Amt
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          
          If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
            Set TempSaleReturn = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & RECEIPT1_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
          Else
            Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & RETURN_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
          End If
          Set HeadCf = m_Headers1.Fields(5)
          If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            SumItem = TempBillingDoc.TOTAL_AMOUNT + TempSaleReturn.TOTAL_AMOUNT
          Else
            SumItem = -TempBillingDoc.TOTAL_AMOUNT
          End If
          TempStr = FormatNumberToNull(SumItem)
          Total1(5) = Total1(5) + SumItem
          Total2(5) = Total2(5) + SumItem
          Total3(5) = Total3(5) + SumItem
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Amt = Amt + TempSaleReturn.TOTAL_PRICE - TempSaleReturn.DISCOUNT_AMOUNT - TempSaleReturn.EXT_DISCOUNT_AMOUNT
          Else
            Amt = -(TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
          End If
          Set HeadCf = m_Headers1.Fields(6)
          TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
          Total1(6) = Total1(6) + MyDiffEx(Amt, SumItem)
          Total2(6) = Total2(6) + MyDiffEx(Amt, SumItem)
          Total3(6) = Total3(6) + MyDiffEx(Amt, SumItem)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set HeadCf = m_Headers1.Fields(7)
          TempStr = FormatNumberToNull(Amt)
          Total1(7) = Total1(7) + Amt
          Total2(7) = Total2(7) + Amt
          Total3(7) = Total3(7) + Amt
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set HeadCf = m_Headers1.Fields(8)
          TempStr = FormatNumberToNull(SumItem)
          Total1(8) = Total1(8) + SumItem
          Total2(8) = Total2(8) + SumItem
          Total3(8) = Total3(8) + SumItem
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            CostAmount = TempBillingDoc.TOTAL_INCLUDE_PRICE
            CostAmount = CostAmount + (TempSaleReturn.TOTAL_INCLUDE_PRICE)
         Else
            CostAmount = -(TempBillingDoc.TOTAL_INCLUDE_PRICE)
         End If
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumberToNull(MyDiffEx(CostAmount, SumItem))
         Total1(9) = Total1(9) + MyDiffEx(CostAmount, SumItem)
         Total2(9) = Total2(9) + MyDiffEx(CostAmount, SumItem)
         Total3(9) = Total3(9) + MyDiffEx(CostAmount, SumItem)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumberToNull(CostAmount)
         Total1(10) = Total1(10) + CostAmount
         Total2(10) = Total2(10) + CostAmount
         Total3(10) = Total3(10) + CostAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          
          Set HeadCf = m_Headers1.Fields(11)
          TempStr = FormatNumberToNull(SumItem)
          Total1(11) = Total1(11) + SumItem
          Total2(11) = Total2(11) + SumItem
          Total3(11) = Total3(11) + SumItem
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set HeadCf = m_Headers1.Fields(12)
          TempStr = FormatNumberToNull(MyDiffEx(Amt - CostAmount, SumItem))
          Total1(12) = Total1(12) + MyDiffEx(Amt - CostAmount, SumItem)
          Total2(12) = Total2(12) + MyDiffEx(Amt - CostAmount, SumItem)
          Total3(12) = Total3(12) + MyDiffEx(Amt - CostAmount, SumItem)
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
          Set HeadCf = m_Headers1.Fields(13)
          TempStr = FormatNumberToNull(Amt - CostAmount)
          Total1(13) = Total1(13) + Amt - CostAmount
          Total2(13) = Total2(13) + Amt - CostAmount
          Total3(13) = Total3(13) + Amt - CostAmount
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
          
         If mcolParam("SHOW_DETAIL") = 1 And (m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or mcolParam("INCLUDE_RETURN") = 1) Then
            Vsp.TextColor = RGB(0, 0, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      End If
   Next m_BillingDoc
   
   If PrevKey2 <> "" Then
      Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      If mcolParam("SHOW_DETAIL") = 1 Then
         Vsp.FontBold = True
         Vsp.TextColor = RGB(0, 0, 255)
      End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      Vsp.TextColor = RGB(0, 0, 0)
   End If
   
   If PrevKey1 <> "" Then
      Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Vsp.TextColor = RGB(0, 0, 255)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      Vsp.TextColor = RGB(0, 0, 0)
   End If
   
   I = 0
   k = 0
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total4(j) = 0
   Next j
   
   For Each m_BillingDoc In DistinctStockNos
      Call CheckNotStock(m_BillingDoc.STOCK_NO, CheckStock)
      If CheckStock = True Then

      k = k + 1
      If ((mcolParam("STOCK_TYPE") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE1") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE2") <> m_BillingDoc.STOCK_TYPE)) Then
         If (PrevKey2 <> m_BillingDoc.STOCK_GROUP_NAME) And (I > 0) Then
            If mcolParam("SUMMARY_NET") = 0 Then
               Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.TextColor = RGB(0, 0, 255)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               Vsp.TextColor = RGB(0, 0, 0)
               
               Call GenerateFooter("--> รวม " & PrevKey2, Total4, m_Details1)
               Vsp.FontBold = True
               Vsp.TextColor = RGB(0, 0, 255)
            Else
               Call GenerateFooter(PrevKey2, Total4, m_Details1)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            Vsp.TextColor = RGB(0, 0, 0)
            
            For j = 1 To UBound(Total1)
               Total2(j) = 0
               Total4(j) = 0
            Next j
            If mcolParam("SUMMARY_NET") = 0 Then
               Vsp = "กลุ่ม :    " & m_BillingDoc.STOCK_GROUP_NAME
               Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
            End If
         ElseIf (PrevKey1 <> m_BillingDoc.STOCK_TYPE_NAME) And (I > 0) Then
            If mcolParam("SUMMARY_NET") = 0 Then
               Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Vsp.TextColor = RGB(0, 0, 255)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               Vsp.TextColor = RGB(0, 0, 0)
            End If
            For j = 1 To UBound(Total1)
               Total2(j) = 0
            Next j
            If mcolParam("SUMMARY_NET") = 0 Then
               Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
            End If
         ElseIf I = 0 Then
            If mcolParam("SUMMARY_NET") = 0 Then
               Vsp = "กลุ่ม :    " & m_BillingDoc.STOCK_GROUP_NAME
               Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
            End If
         End If
             
         I = I + 1
         Call m_Details1.ClearField
         Vsp.TextColor = RGB(0, 0, 0)
         
         PrevKey1 = m_BillingDoc.STOCK_TYPE_NAME
         PrevKey2 = m_BillingDoc.STOCK_GROUP_NAME
         
            Debug.Print m_BillingDoc.STOCK_NO
         
        If m_BillingDoc.STOCK_NO = "20-1006" Then
            Debug.Print
         End If
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_BillingDoc.STOCK_DESC & " /" & m_BillingDoc.STOCK_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set TempBillingDoc = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE), True)
         Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE), True)
         Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE), True)
         
         SumItem = TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = FormatNumberToNull(SumItem)
         Total2(2) = Total2(2) + SumItem
         Total3(2) = Total3(2) + SumItem
         Total4(2) = Total4(2) + SumItem
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
         Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
         
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
         Total2(3) = Total2(3) + MyDiffEx(Amt, SumItem)
         Total3(3) = Total3(3) + MyDiffEx(Amt, SumItem)
         Total4(3) = Total4(3) + MyDiffEx(Amt, SumItem)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = FormatNumberToNull(Amt)
         Total2(4) = Total2(4) + Amt
         Total3(4) = Total3(4) + Amt
         Total4(4) = Total4(4) + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set TempBillingDoc = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE), True)
         Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE), True)
         Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE), True)
         
         SumItem = TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
         
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = FormatNumberToNull(SumItem)
         Total2(5) = Total2(5) + SumItem
         Total3(5) = Total3(5) + SumItem
         Total4(5) = Total4(5) + SumItem
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
         Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
          
         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
         Total2(6) = Total2(6) + MyDiffEx(Amt, SumItem)
         Total3(6) = Total3(6) + MyDiffEx(Amt, SumItem)
         Total4(6) = Total4(6) + MyDiffEx(Amt, SumItem)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumberToNull(Amt)
         Total2(7) = Total2(7) + Amt
         Total3(7) = Total3(7) + Amt
         Total4(7) = Total4(7) + Amt
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = FormatNumberToNull(SumItem)
         Total2(8) = Total2(8) + SumItem
         Total3(8) = Total3(8) + SumItem
         Total4(8) = Total4(8) + SumItem
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         CostAmount = (TempBillingDoc.TOTAL_INCLUDE_PRICE) + (TempBillingDoc2.TOTAL_INCLUDE_PRICE) - (TempBillingDoc3.TOTAL_INCLUDE_PRICE)
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumberToNull(MyDiffEx(CostAmount, SumItem))
         Total2(9) = Total2(9) + MyDiffEx(CostAmount, SumItem)
         Total3(9) = Total3(9) + MyDiffEx(CostAmount, SumItem)
         Total4(9) = Total4(9) + MyDiffEx(CostAmount, SumItem)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(10)
         TempStr = FormatNumberToNull(CostAmount)
         Total2(10) = Total2(10) + CostAmount
         Total3(10) = Total3(10) + CostAmount
         Total4(10) = Total4(10) + CostAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(11)
         TempStr = FormatNumberToNull(SumItem)
         Total2(11) = Total2(11) + SumItem
         Total3(11) = Total3(11) + SumItem
         Total4(11) = Total4(11) + SumItem
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(12)
         TempStr = FormatNumberToNull(MyDiffEx(Amt - CostAmount, SumItem))
         Total2(12) = Total2(12) + MyDiffEx(Amt - CostAmount, SumItem)
         Total3(12) = Total3(12) + MyDiffEx(Amt - CostAmount, SumItem)
         Total4(12) = Total4(12) + MyDiffEx(Amt - CostAmount, SumItem)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(13)
         TempStr = FormatNumberToNull(Amt - CostAmount)
         Total2(13) = Total2(13) + Amt - CostAmount
         Total3(13) = Total3(13) + Amt - CostAmount
         Total4(13) = Total4(13) + Amt - CostAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If mcolParam("SUMMARY") = 0 And mcolParam("SUMMARY_NET") = 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      Else
         Set TempBillingDoc = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE), True)
         Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE), True)
         Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE), True)
         
         SumItem = TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
         Total4(2) = Total4(2) + SumItem
          
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
         Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
         Total4(3) = Total4(3) + MyDiffEx(Amt, SumItem)
         Total4(4) = Total4(4) + Amt
         
         Set TempBillingDoc = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE), True)
         Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE), True)
         Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE), True)
         
         SumItem = TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
         Total4(5) = Total4(5) + SumItem
          
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
         Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
          
         Total4(6) = Total4(6) + MyDiffEx(Amt, SumItem)
         Total4(7) = Total4(7) + Amt
         Total4(8) = Total4(8) + SumItem
         CostAmount = (TempBillingDoc.TOTAL_INCLUDE_PRICE) + (TempBillingDoc2.TOTAL_INCLUDE_PRICE) - (TempBillingDoc3.TOTAL_INCLUDE_PRICE)
         Total4(9) = Total4(9) + MyDiffEx(CostAmount, SumItem)
         Total4(10) = Total4(10) + CostAmount
         Total4(11) = Total4(11) + SumItem
         Total4(12) = Total4(12) + MyDiffEx(Amt - CostAmount, SumItem)
         Total4(13) = Total4(13) + Amt - CostAmount
         
      End If
      End If
   Next m_BillingDoc
   
   If (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
      If mcolParam("SUMMARY_NET") = 0 Then
         Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Vsp.TextColor = RGB(0, 0, 255)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         Vsp.TextColor = RGB(0, 0, 0)
         
         Call GenerateFooter("--> รวม " & PrevKey2, Total4, m_Details1)
         Vsp.FontBold = True
         Vsp.TextColor = RGB(0, 0, 255)
      Else
         Call GenerateFooter(PrevKey2, Total4, m_Details1)
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      Vsp.TextColor = RGB(0, 0, 0)
   End If
   
   Dim TempX As Double
   Dim TempY As Double
   
   TempX = Vsp.CurrentX
   TempY = Vsp.CurrentY
   
   Call GenerateFooterEnd("                        สุทธิ", Total3, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TextColor = RGB(0, 0, 255)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False

   If mcolParam("CONSIGNMENT") = 1 Then
      Vsp = ""
      Vsp = "ฝากขาย"
      Vsp.TextColor = RGB(0, 0, 0)
      I = 0
      k = 0
      For j = 1 To UBound(Total1)
         Total1(j) = 0
         Total2(j) = 0
         Total3(j) = 0
         Total4(j) = 0
      Next j
   
      For Each m_BillingDoc In DistinctStockNosConsignment
         Call CheckNotStock(m_BillingDoc.STOCK_NO, CheckStock)
         If CheckStock = True Then
            k = k + 1
            If ((mcolParam("STOCK_TYPE") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE1") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE2") <> m_BillingDoc.STOCK_TYPE)) Then
               If (PrevKey2 <> m_BillingDoc.STOCK_GROUP_NAME) And (I > 0) Then
                  If mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
                     Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Vsp.TextColor = RGB(0, 0, 255)
                     Vsp.FontBold = True
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                     Vsp.TextColor = RGB(0, 0, 0)
                     
                     Call GenerateFooter("--> รวม " & PrevKey2, Total4, m_Details1)
                     Vsp.FontBold = True
                     Vsp.TextColor = RGB(0, 0, 255)
                  Else
                     Call GenerateFooter(PrevKey2, Total4, m_Details1)
                  End If
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  Vsp.TextColor = RGB(0, 0, 0)
                  
                  For j = 1 To UBound(Total1)
                     Total2(j) = 0
                     Total4(j) = 0
                  Next j
                  If mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
                     Vsp = "กลุ่ม :    " & m_BillingDoc.STOCK_GROUP_NAME
                     Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
                  End If
               ElseIf (PrevKey1 <> m_BillingDoc.STOCK_TYPE_NAME) And (I > 0) Then
                  If mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
                     Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
                     Call m_Details1.GetString(1, TempStr1, TempStr2)
                     strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                     Vsp.FontBold = True
                     Vsp.TextColor = RGB(0, 0, 255)
                     Call Vsp.AddTable(strFormat, "", TempStr2)
                     Vsp.FontBold = False
                     Vsp.TextColor = RGB(0, 0, 0)
                  End If
                  For j = 1 To UBound(Total1)
                     Total2(j) = 0
                  Next j
                  If mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
                     Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
                  End If
               ElseIf I = 0 Then
                  If mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
                     Vsp = "กลุ่ม :    " & m_BillingDoc.STOCK_GROUP_NAME
                     Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
                  End If
               End If
                   
               I = I + 1
               Call m_Details1.ClearField
               Vsp.TextColor = RGB(0, 0, 0)
               
               PrevKey1 = m_BillingDoc.STOCK_TYPE_NAME
               PrevKey2 = m_BillingDoc.STOCK_GROUP_NAME
               
               Set HeadCf = m_Headers1.Fields(1)
               TempStr = m_BillingDoc.STOCK_DESC & " /" & m_BillingDoc.STOCK_NO
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set TempBillingDoc = GetObject("CBillingDoc", ConsignmentCmps, Trim(m_BillingDoc.STOCK_NO & "-" & PO_DOCTYPE), True)
               SumItem = TempBillingDoc.TOTAL_AMOUNT
               
               Set HeadCf = m_Headers1.Fields(2)
               TempStr = FormatNumberToNull(SumItem)
               Total2(2) = Total2(2) + SumItem
               Total3(2) = Total3(2) + SumItem
               Total4(2) = Total4(2) + SumItem
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                
               Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               
               Set HeadCf = m_Headers1.Fields(3)
               TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
               Total2(3) = Total2(3) + MyDiffEx(Amt, SumItem)
               Total3(3) = Total3(3) + MyDiffEx(Amt, SumItem)
               Total4(3) = Total4(3) + MyDiffEx(Amt, SumItem)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(4)
               TempStr = FormatNumberToNull(Amt)
               Total2(4) = Total2(4) + Amt
               Total3(4) = Total3(4) + Amt
               Total4(4) = Total4(4) + Amt
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set TempBillingDoc = GetObject("CBillingDoc", CONSIGNMENT, Trim(m_BillingDoc.STOCK_NO & "-" & PO_DOCTYPE), True)
               
               SumItem = TempBillingDoc.TOTAL_AMOUNT
               
               Set HeadCf = m_Headers1.Fields(5)
               TempStr = FormatNumberToNull(SumItem)
               Total2(5) = Total2(5) + SumItem
               Total3(5) = Total3(5) + SumItem
               Total4(5) = Total4(5) + SumItem
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
                
               Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
                
               Set HeadCf = m_Headers1.Fields(6)
               TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
               Total2(6) = Total2(6) + MyDiffEx(Amt, SumItem)
               Total3(6) = Total3(6) + MyDiffEx(Amt, SumItem)
               Total4(6) = Total4(6) + MyDiffEx(Amt, SumItem)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(7)
               TempStr = FormatNumberToNull(Amt)
               Total2(7) = Total2(7) + Amt
               Total3(7) = Total3(7) + Amt
               Total4(7) = Total4(7) + Amt
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(8)
               TempStr = FormatNumberToNull(SumItem)
               Total2(8) = Total2(8) + SumItem
               Total3(8) = Total3(8) + SumItem
               Total4(8) = Total4(8) + SumItem
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               CostAmount = (TempBillingDoc.TOTAL_INCLUDE_PRICE)
               
               Set HeadCf = m_Headers1.Fields(9)
               TempStr = FormatNumberToNull(MyDiffEx(CostAmount, SumItem))
               Total2(9) = Total2(9) + MyDiffEx(CostAmount, SumItem)
               Total3(9) = Total3(9) + MyDiffEx(CostAmount, SumItem)
               Total4(9) = Total4(9) + MyDiffEx(CostAmount, SumItem)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(10)
               TempStr = FormatNumberToNull(CostAmount)
               Total2(10) = Total2(10) + CostAmount
               Total3(10) = Total3(10) + CostAmount
               Total4(10) = Total4(10) + CostAmount
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(11)
               TempStr = FormatNumberToNull(SumItem)
               Total2(11) = Total2(11) + SumItem
               Total3(11) = Total3(11) + SumItem
               Total4(11) = Total4(11) + SumItem
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(12)
               TempStr = FormatNumberToNull(MyDiffEx(Amt - CostAmount, SumItem))
               Total2(12) = Total2(12) + MyDiffEx(Amt - CostAmount, SumItem)
               Total3(12) = Total3(12) + MyDiffEx(Amt - CostAmount, SumItem)
               Total4(12) = Total4(12) + MyDiffEx(Amt - CostAmount, SumItem)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               Set HeadCf = m_Headers1.Fields(13)
               TempStr = FormatNumberToNull(Amt - CostAmount)
               Total2(13) = Total2(13) + Amt - CostAmount
               Total3(13) = Total3(13) + Amt - CostAmount
               Total4(13) = Total4(13) + Amt - CostAmount
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               
               If mcolParam("SUMMARY_CONSIGNMENT") = 0 And mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
               End If
            Else
               Set TempBillingDoc = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & PO_DOCTYPE), True)
               
               SumItem = TempBillingDoc.TOTAL_AMOUNT
               Total4(2) = Total4(2) + SumItem
                
               Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Total4(3) = Total4(3) + MyDiffEx(Amt, SumItem)
               Total4(4) = Total4(4) + Amt
               
               Set TempBillingDoc = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & PO_DOCTYPE), True)
               
               SumItem = TempBillingDoc.TOTAL_AMOUNT
               Total4(5) = Total4(5) + SumItem
                
               Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Total4(6) = Total4(6) + MyDiffEx(Amt, SumItem)
               Total4(7) = Total4(7) + Amt
               Total4(8) = Total4(8) + SumItem
               CostAmount = (TempBillingDoc.TOTAL_INCLUDE_PRICE)
               
               Total4(9) = Total4(9) + MyDiffEx(CostAmount, SumItem)
               Total4(10) = Total4(10) + CostAmount
               Total4(11) = Total4(11) + SumItem
               Total4(12) = Total4(12) + MyDiffEx(Amt - CostAmount, SumItem)
               Total4(13) = Total4(13) + Amt - CostAmount
            End If
         End If
      Next m_BillingDoc
      
      If (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
           If mcolParam("SUMMARY_CONSIGNMENT_NET") = 0 Then
              Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
              Call m_Details1.GetString(1, TempStr1, TempStr2)
              strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
              Vsp.FontBold = True
              Vsp.TextColor = RGB(0, 0, 255)
              Call Vsp.AddTable(strFormat, "", TempStr2)
              Vsp.FontBold = False
              Vsp.TextColor = RGB(0, 0, 0)
              
              Call GenerateFooter("--> รวม " & PrevKey2, Total4, m_Details1)
              Vsp.FontBold = True
              Vsp.TextColor = RGB(0, 0, 255)
           Else
              Call GenerateFooter(PrevKey2, Total4, m_Details1)
           End If
           Call m_Details1.GetString(1, TempStr1, TempStr2)
           strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
           Call Vsp.AddTable(strFormat, "", TempStr2)
           Vsp.FontBold = False
           Vsp.TextColor = RGB(0, 0, 0)
        End If

        TempX = Vsp.CurrentX
        TempY = Vsp.CurrentY
        
        Call GenerateFooterEnd("                        สุทธิ", Total3, m_Details1)
        Call m_Details1.GetString(1, TempStr1, TempStr2)
        strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
        Vsp.FontBold = True
        Vsp.TextColor = RGB(0, 0, 255)
        Call Vsp.AddTable(strFormat, "", TempStr2)
        Vsp.FontBold = False
   End If

   Call GenerateEndDoc
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set m_BillingDoc = Nothing
   Set TempBillingDoc = Nothing
   Set TotalMoves = Nothing
   Set SaleReturnColl = Nothing
   MasterInd = "1"
   
   genDoc = True
   Exit Function
   
ErrHandler:
   MasterInd = "1"
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempAmount As Double
Dim TempSale As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j))
      Else
         TempStr = FormatNumberToNull(MyDiff(Tot(j + 1), Tot(j - 1)))
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooterEnd(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempAmount As Double
Dim TempSale As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j))
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
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub
Private Sub CheckNotStock(StockNo As String, CheckStock As Boolean)
Dim TempStock As CStockCode
      If mcolParam("NOT_FROM_STOCK_NO") <> "" Or mcolParam("NOT_TO_STOCK_NO") <> "" Then
         Set TempStock = GetObject("CBillingDoc", NotStockNo, Trim(StockNo), False)
         If TempStock Is Nothing Then
            CheckStock = True
         Else
            CheckStock = False
         End If
      Else
         CheckStock = True
      End If
End Sub
