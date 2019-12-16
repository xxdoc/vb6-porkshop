VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling035_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling035_1"
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

Private DateCount As Long

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
   Set DistinctStockNos = New Collection
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
   Set DistinctStockNos = Nothing
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
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim Amt As Double
Dim iCount As Long
Dim TempDateString As String
Dim DateCount  As Long
Dim TempFromDate  As Date
Dim TempToDate   As Date
Dim CountFlag  As Double
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   
   CountFlag = 0
      
   If mcolParam("SHOW_AMOUNT") > 0 Then
      CountFlag = CountFlag + 2
   End If
   If mcolParam("SHOW_PRICE") > 0 Then
      CountFlag = CountFlag + 2
   End If
   If mcolParam("SHOW_COST") > 0 Then
      CountFlag = CountFlag + 2
   End If
   If mcolParam("SHOW_GP") > 0 Then
      CountFlag = CountFlag + 2
   End If
      
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   
   Call Cf.SetFieldValue(5, "^", "รายการเอียด", "<")
   Call m_Headers0.AddField(Cf)
   TempFromDate = mcolParam("FROM_BILL_DATE")
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
      TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)
      
      Call Cf.SetFieldValue(CountFlag, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
      Call m_Headers0.AddField(Cf)
      
      TempFromDate = DateAdd("D", 1, TempToDate)
   Wend
   
   Call Cf.SetFieldValue(CountFlag, "^", "รวม", ">")
   Call m_Headers0.AddField(Cf)
   
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   
   Call Cf.SetFieldValue(5, "^", "รายการสินค้า/รหัส", "<")
   Call m_Headers1.AddField(Cf)
   
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      If mcolParam("SHOW_AMOUNT") > 0 Then
         CountFlag = 2
      Else
         CountFlag = 0
      End If
      Call Cf.SetFieldValue(CountFlag, "^", "จำนวน", ">", , True)
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      If mcolParam("SHOW_PRICE") > 0 Then
         CountFlag = 2
      Else
         CountFlag = 0
      End If
      Call Cf.SetFieldValue(CountFlag, "^", "ยอดเงิน", ">", , True)
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      If mcolParam("SHOW_COST") > 0 Then
         CountFlag = 2
      Else
         CountFlag = 0
      End If
      Call Cf.SetFieldValue(CountFlag, "^", "ต้นทุน", ">", , True)
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(0, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      If mcolParam("SHOW_GP") > 0 Then
         CountFlag = 2
      Else
         CountFlag = 0
      End If
      Call Cf.SetFieldValue(CountFlag, "^", "กำไร", ">", , True)
      Call m_Headers1.AddField(Cf)
   Wend
   
   If mcolParam("SHOW_AMOUNT") > 0 Then
      CountFlag = 2
   Else
      CountFlag = 0
   End If
   Call Cf.SetFieldValue(CountFlag, "^", "จำนวน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_PRICE") > 0 Then
      CountFlag = 2
   Else
      CountFlag = 0
   End If
   Call Cf.SetFieldValue(CountFlag, "^", "ยอดเงิน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_COST") > 0 Then
      CountFlag = 2
   Else
      CountFlag = 0
   End If
   Call Cf.SetFieldValue(CountFlag, "^", "ต้นทุน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(0, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_GP") > 0 Then
      CountFlag = 2
   Else
      CountFlag = 0
   End If
   Call Cf.SetFieldValue(CountFlag, "^", "กำไร", ">", , True)
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
Dim FromDate As String
Dim ToDate As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orLandscape
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
         
   Call GetSaleAmountAparGroupStockCodeDateFree(TotalMoves, mcolParam("FROM_BILL_DATE"), DateAdd("D", DateCount - 1, mcolParam("FROM_BILL_DATE")), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   
   If (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
      Call GetDistinctStockCode(DistinctStockNos, mcolParam("FROM_BILL_DATE"), DateAdd("D", DateCount - 1, mcolParam("FROM_BILL_DATE")), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
      Call GetSaleAmountStockCodeDocTypeDate(OtherMoves, mcolParam("FROM_BILL_DATE"), DateAdd("D", DateCount - 1, mcolParam("FROM_BILL_DATE")), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   End If
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(DateAdd("D", DateCount - 1, mcolParam("FROM_BILL_DATE"))), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A")
                                       
                                       
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
Dim Total0(100) As Double
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim SumItem As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String

Dim CostAmount As Double

Dim SaleReturnColl As Collection
Dim SaleReturnArrangeColl  As Collection
Dim TempSaleReturn As CBillingDoc
Dim TempDateString As String
Dim CountDate As Long
Dim TempDate As Date

Dim TotalAmount As Double
Dim TotalPrice As Double
Dim TotalCost As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set SaleReturnColl = New Collection
   Set SaleReturnArrangeColl = New Collection
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
      Total0(j) = 0
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
   
   MasterInd = "92"
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
   m_BillingDoc.TO_DATE = DateAdd("D", DateCount - 1, mcolParam("FROM_BILL_DATE"))
   m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.FREE_FLAG = StringToFreeFlag(mcolParam("INCLUDE_FREE"))
   Call m_BillingDoc.QueryData(92, Rs, iCount)
   
   I = 0
   
   While Not Rs.EOF
      MasterInd = "92"
      Set m_BillingDoc = New CBillingDoc
      
      Call m_BillingDoc.PopulateFromRS2(92, Rs)
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
      
      Rs.MoveNext
   Wend
   
   PrevKey1 = ""
   For Each m_BillingDoc In SaleReturnColl    'จัดเรียง
      If PrevKey1 <> Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & INVOICE_DOCTYPE) And PrevKey1 <> "" Then
         For Each TempSaleReturn In SaleReturnColl
            If Trim(TempSaleReturn.APAR_GROUP_NAME & "-" & TempSaleReturn.DOCUMENT_TYPE) = PrevKey1 Then
               Set TempBillingDoc = GetObject("CBillingDoc", SaleReturnArrangeColl, Trim(TempSaleReturn.APAR_GROUP_NAME & "-" & TempSaleReturn.DOCUMENT_TYPE & "-" & TempSaleReturn.STOCK_NO), False)
               If TempBillingDoc Is Nothing Then
                  Set TempBillingDoc = New CBillingDoc
                  TempBillingDoc.APAR_GROUP_NAME = TempSaleReturn.APAR_GROUP_NAME
                  TempBillingDoc.STOCK_TYPE_NAME = TempSaleReturn.STOCK_TYPE_NAME
                  TempBillingDoc.STOCK_NO = TempSaleReturn.STOCK_NO
                  TempBillingDoc.STOCK_DESC = TempSaleReturn.STOCK_DESC
                  TempBillingDoc.STOCK_TYPE = TempSaleReturn.STOCK_TYPE
                  TempBillingDoc.CHK_STD_COST = TempSaleReturn.CHK_STD_COST
                  TempBillingDoc.COST_PER_AMOUNT = TempSaleReturn.COST_PER_AMOUNT
                  TempBillingDoc.DOCUMENT_TYPE = TempSaleReturn.DOCUMENT_TYPE
         
                  Call SaleReturnArrangeColl.add(TempBillingDoc, Trim(TempSaleReturn.APAR_GROUP_NAME & "-" & TempSaleReturn.DOCUMENT_TYPE & "-" & TempSaleReturn.STOCK_NO))
               End If
            End If
         Next TempSaleReturn
      End If
      
      Set TempBillingDoc = GetObject("CBillingDoc", SaleReturnArrangeColl, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & m_BillingDoc.DOCUMENT_TYPE & "-" & m_BillingDoc.STOCK_NO), False)
      If TempBillingDoc Is Nothing Then
         Set TempBillingDoc = New CBillingDoc
         TempBillingDoc.APAR_GROUP_NAME = m_BillingDoc.APAR_GROUP_NAME
         TempBillingDoc.STOCK_TYPE_NAME = m_BillingDoc.STOCK_TYPE_NAME
         TempBillingDoc.STOCK_NO = m_BillingDoc.STOCK_NO
         TempBillingDoc.STOCK_DESC = m_BillingDoc.STOCK_DESC
         TempBillingDoc.STOCK_TYPE = m_BillingDoc.STOCK_TYPE
         TempBillingDoc.CHK_STD_COST = m_BillingDoc.CHK_STD_COST
         TempBillingDoc.COST_PER_AMOUNT = m_BillingDoc.COST_PER_AMOUNT
         TempBillingDoc.DOCUMENT_TYPE = m_BillingDoc.DOCUMENT_TYPE
         
         Call SaleReturnArrangeColl.add(TempBillingDoc, Trim(TempBillingDoc.APAR_GROUP_NAME & "-" & TempBillingDoc.DOCUMENT_TYPE & "-" & TempBillingDoc.STOCK_NO))
         PrevKey1 = Trim(TempBillingDoc.APAR_GROUP_NAME & "-" & TempBillingDoc.DOCUMENT_TYPE)
      End If
   Next m_BillingDoc
   
   I = 0
   PrevKey1 = ""
   PrevKey2 = ""
   For Each m_BillingDoc In SaleReturnArrangeColl
      I = I + 1
      Call m_Details1.ClearField
      If ((mcolParam("STOCK_TYPE") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE1") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE2") <> m_BillingDoc.STOCK_TYPE)) And (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
         I = I - 1
      Else
         If PrevKey1 <> m_BillingDoc.APAR_GROUP_NAME And I <> 1 Then
             
            Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
            
            Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
             
            
             Vsp = "กลุ่มลูกค้า  :  " & m_BillingDoc.APAR_GROUP_NAME
             ' รวม
             For j = 1 To UBound(Total1)
                Total1(j) = 0
                Total2(j) = 0
             Next j
             
          ElseIf PrevKey2 <> Trim(Str(m_BillingDoc.DOCUMENT_TYPE)) And I <> 1 Then
             Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
             Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Vsp.FontBold = True
             Call Vsp.AddTable(strFormat, "", TempStr2)
             Vsp.FontBold = False
            
            ' รวม
             For j = 1 To UBound(Total1)
                Total1(j) = 0
             Next j
         ElseIf I = 1 Then
             Vsp = "กลุ่มลูกค้า  :  " & m_BillingDoc.APAR_GROUP_NAME
          End If
          
          PrevKey1 = m_BillingDoc.APAR_GROUP_NAME
          PrevKey2 = Trim(Str(m_BillingDoc.DOCUMENT_TYPE))
          
          Call m_Details1.ClearField
               
         k = 0
         
         k = k + 1
          Set HeadCf = m_Headers1.Fields(k)
          TempStr = m_BillingDoc.STOCK_DESC & " /" & m_BillingDoc.STOCK_NO
          Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
          Call m_Details1.AddField(BodyCf)
                   
          TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
         TempDateString = Trim(Replace(TempDateString, ")", ""))
         TempDate = DateAdd("D", -1, mcolParam("FROM_BILL_DATE"))
         
         TotalAmount = 0
         TotalPrice = 0
         TotalCost = 0
         
         While Len(TempDateString) > 0
          Call CalculateDatePeriod(TempDateString, CountDate)
          
          SumItem = 0
          Amt = 0
          CostAmount = 0
          While CountDate > 0
               TempDate = DateAdd("D", 1, TempDate)
              If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
                 Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO & "-" & TempDate), True)
                 Set TempSaleReturn = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & RECEIPT1_DOCTYPE & "-" & m_BillingDoc.STOCK_NO & "-" & TempDate), True)
              Else
                 Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_GROUP_NAME & "-" & RETURN_DOCTYPE & "-" & m_BillingDoc.STOCK_NO & "-" & TempDate), True)
              End If
              
              If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
                 SumItem = SumItem + (TempBillingDoc.TOTAL_AMOUNT + TempSaleReturn.TOTAL_AMOUNT)
              Else
                 SumItem = SumItem - TempBillingDoc.TOTAL_AMOUNT
              End If
            
              If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
                 Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
                 Amt = Amt + TempSaleReturn.TOTAL_PRICE - TempSaleReturn.DISCOUNT_AMOUNT - TempSaleReturn.EXT_DISCOUNT_AMOUNT
              Else
                 Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
              End If
              
              If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
                  CostAmount = CostAmount + (TempBillingDoc.TOTAL_INCLUDE_PRICE)
                  CostAmount = CostAmount + (TempSaleReturn.TOTAL_INCLUDE_PRICE)
               Else
                  CostAmount = CostAmount - (TempBillingDoc.TOTAL_INCLUDE_PRICE)
               End If
              CountDate = CountDate - 1
            Wend
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(SumItem)
            Total0(k) = Total0(k) + SumItem
            Total1(k) = Total1(k) + SumItem
            Total2(k) = Total2(k) + SumItem
            Total3(k) = Total3(k) + SumItem
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
          
             k = k + 1
             Set HeadCf = m_Headers1.Fields(k)
             TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
             Total0(k) = Total0(k) + MyDiffEx(Amt, SumItem)
             Total1(k) = Total1(k) + MyDiffEx(Amt, SumItem)
             Total2(k) = Total2(k) + MyDiffEx(Amt, SumItem)
             Total3(k) = Total3(k) + MyDiffEx(Amt, SumItem)
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
             Call m_Details1.AddField(BodyCf)
             
             k = k + 1
             Set HeadCf = m_Headers1.Fields(k)
             TempStr = FormatNumberToNull(Amt)
             Total0(k) = Total0(k) + Amt
             Total1(k) = Total1(k) + Amt
             Total2(k) = Total2(k) + Amt
             Total3(k) = Total3(k) + Amt
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
             Call m_Details1.AddField(BodyCf)
             
             k = k + 1
             Set HeadCf = m_Headers1.Fields(k)
             TempStr = FormatNumberToNull(SumItem)
             Total0(k) = Total0(k) + SumItem
             Total1(k) = Total1(k) + SumItem
             Total2(k) = Total2(k) + SumItem
             Total3(k) = Total3(k) + SumItem
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
             Call m_Details1.AddField(BodyCf)
             
             k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(MyDiffEx(CostAmount, SumItem))
            Total0(k) = Total0(k) + MyDiffEx(CostAmount, SumItem)
            Total1(k) = Total1(k) + MyDiffEx(CostAmount, SumItem)
            Total2(k) = Total2(k) + MyDiffEx(CostAmount, SumItem)
            Total3(k) = Total3(k) + MyDiffEx(CostAmount, SumItem)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
             
             k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(CostAmount)
            Total0(k) = Total0(k) + CostAmount
            Total1(k) = Total1(k) + CostAmount
            Total2(k) = Total2(k) + CostAmount
            Total3(k) = Total3(k) + CostAmount
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
             
             k = k + 1
             Set HeadCf = m_Headers1.Fields(k)
             TempStr = FormatNumberToNull(SumItem)
             Total0(k) = Total0(k) + SumItem
             Total1(k) = Total1(k) + SumItem
             Total2(k) = Total2(k) + SumItem
             Total3(k) = Total3(k) + SumItem
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
             Call m_Details1.AddField(BodyCf)
             
             k = k + 1
             Set HeadCf = m_Headers1.Fields(k)
             TempStr = FormatNumberToNull(MyDiffEx(Amt - CostAmount, SumItem))
             Total0(k) = Total0(k) + MyDiffEx(Amt - CostAmount, SumItem)
             Total1(k) = Total1(k) + MyDiffEx(Amt - CostAmount, SumItem)
             Total2(k) = Total2(k) + MyDiffEx(Amt - CostAmount, SumItem)
             Total3(k) = Total3(k) + MyDiffEx(Amt - CostAmount, SumItem)
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
             Call m_Details1.AddField(BodyCf)
             
             k = k + 1
             Set HeadCf = m_Headers1.Fields(k)
             TempStr = FormatNumberToNull(Amt - CostAmount)
             Total0(k) = Total0(k) + Amt - CostAmount
             Total1(k) = Total1(k) + Amt - CostAmount
             Total2(k) = Total2(k) + Amt - CostAmount
             Total3(k) = Total3(k) + Amt - CostAmount
             Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
             Call m_Details1.AddField(BodyCf)
             
             TotalAmount = TotalAmount + SumItem
             TotalPrice = TotalPrice + Amt
             TotalCost = TotalCost + CostAmount
          Wend
          
         k = k + 1
        Set HeadCf = m_Headers1.Fields(k)
        TempStr = FormatNumberToNull(TotalAmount)
        Total0(k) = Total0(k) + TotalAmount
        Total1(k) = Total1(k) + TotalAmount
        Total2(k) = Total2(k) + TotalAmount
        Total3(k) = Total3(k) + TotalAmount
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
      
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(MyDiffEx(Amt, TotalAmount))
         Total0(k) = Total0(k) + MyDiffEx(TotalPrice, TotalAmount)
         Total1(k) = Total1(k) + MyDiffEx(TotalPrice, TotalAmount)
         Total2(k) = Total2(k) + MyDiffEx(TotalPrice, TotalAmount)
         Total3(k) = Total3(k) + MyDiffEx(TotalPrice, TotalAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalPrice)
         Total0(k) = Total0(k) + TotalPrice
         Total1(k) = Total1(k) + TotalPrice
         Total2(k) = Total2(k) + TotalPrice
         Total3(k) = Total3(k) + TotalPrice
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalAmount)
         Total0(k) = Total0(k) + TotalAmount
         Total1(k) = Total1(k) + TotalAmount
         Total2(k) = Total2(k) + TotalAmount
         Total3(k) = Total3(k) + TotalAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
        Set HeadCf = m_Headers1.Fields(k)
        TempStr = FormatNumberToNull(MyDiffEx(TotalCost, TotalAmount))
        Total0(k) = Total0(k) + MyDiffEx(TotalCost, TotalAmount)
        Total1(k) = Total1(k) + MyDiffEx(TotalCost, TotalAmount)
        Total2(k) = Total2(k) + MyDiffEx(TotalCost, TotalAmount)
        Total3(k) = Total3(k) + MyDiffEx(TotalCost, TotalAmount)
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         k = k + 1
        Set HeadCf = m_Headers1.Fields(k)
        TempStr = FormatNumberToNull(TotalCost)
        Total0(k) = Total0(k) + TotalCost
        Total1(k) = Total1(k) + TotalCost
        Total2(k) = Total2(k) + TotalCost
        Total3(k) = Total3(k) + TotalCost
        Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
        Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalAmount)
         Total0(k) = Total0(k) + TotalAmount
         Total1(k) = Total1(k) + TotalAmount
         Total2(k) = Total2(k) + TotalAmount
         Total3(k) = Total3(k) + TotalAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(MyDiffEx(TotalPrice - TotalCost, TotalAmount))
         Total0(k) = Total0(k) + MyDiffEx(TotalPrice - TotalCost, TotalAmount)
         Total1(k) = Total1(k) + MyDiffEx(TotalPrice - TotalCost, TotalAmount)
         Total2(k) = Total2(k) + MyDiffEx(TotalPrice - TotalCost, TotalAmount)
         Total3(k) = Total3(k) + MyDiffEx(TotalPrice - TotalCost, TotalAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalPrice - TotalCost)
         Total0(k) = Total0(k) + TotalPrice - TotalCost
         Total1(k) = Total1(k) + TotalPrice - TotalCost
         Total2(k) = Total2(k) + TotalPrice - TotalCost
         Total3(k) = Total3(k) + TotalPrice - TotalCost
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If mcolParam("SHOW_DETAIL") = 1 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      End If
   Next m_BillingDoc
   
   Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   I = 0
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
            
   For Each m_BillingDoc In DistinctStockNos
      If ((mcolParam("STOCK_TYPE") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE1") <> m_BillingDoc.STOCK_TYPE) And (mcolParam("STOCK_TYPE2") <> m_BillingDoc.STOCK_TYPE)) Then
         
         If (PrevKey1 <> m_BillingDoc.STOCK_TYPE_NAME) And (I > 0) Then
            Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total2(j) = 0
            Next j
            Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
         ElseIf I = 0 Then
            Vsp = "ประเภท :    " & m_BillingDoc.STOCK_TYPE_NAME
         End If
             
          I = I + 1
          k = 0
         Call m_Details1.ClearField
         
         k = k + 1
         PrevKey1 = m_BillingDoc.STOCK_TYPE_NAME
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = m_BillingDoc.STOCK_DESC & " /" & m_BillingDoc.STOCK_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
          TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
         TempDateString = Trim(Replace(TempDateString, ")", ""))
         TempDate = DateAdd("D", -1, mcolParam("FROM_BILL_DATE"))
         
         TotalAmount = 0
         TotalPrice = 0
         TotalCost = 0
         
         While Len(TempDateString) > 0
          k = k + 1
          Call CalculateDatePeriod(TempDateString, CountDate)

          SumItem = 0
          Amt = 0
          CostAmount = 0
          While CountDate > 0
               TempDate = DateAdd("D", 1, TempDate)
         
               Set TempBillingDoc = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE & "-" & TempDate), True)
               Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE & "-" & TempDate), True)
               Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE & "-" & TempDate), True)
               
               SumItem = SumItem + TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
               
               Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
               Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
                
                CostAmount = CostAmount + (TempBillingDoc.TOTAL_INCLUDE_PRICE) + (TempBillingDoc2.TOTAL_INCLUDE_PRICE) - (TempBillingDoc3.TOTAL_INCLUDE_PRICE)
               
               CountDate = CountDate - 1
            Wend
         
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(SumItem)
            Total2(k) = Total2(k) + SumItem
            Total3(k) = Total3(k) + SumItem
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
             
             k = k + 1
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
            Total2(k) = Total2(k) + MyDiffEx(Amt, SumItem)
            Total3(k) = Total3(k) + MyDiffEx(Amt, SumItem)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt)
            Total2(k) = Total2(k) + Amt
            Total3(k) = Total3(k) + Amt
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
                     
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(SumItem)
            Total2(k) = Total2(k) + SumItem
            Total3(k) = Total3(k) + SumItem
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(MyDiffEx(CostAmount, SumItem))
            Total2(k) = Total2(k) + MyDiffEx(CostAmount, SumItem)
            Total3(k) = Total3(k) + MyDiffEx(CostAmount, SumItem)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(CostAmount)
            Total2(k) = Total2(k) + CostAmount
            Total3(k) = Total3(k) + CostAmount
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(SumItem)
            Total2(k) = Total2(k) + SumItem
            Total3(k) = Total3(k) + SumItem
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(MyDiffEx(Amt - CostAmount, SumItem))
            Total2(k) = Total2(k) + MyDiffEx(Amt - CostAmount, SumItem)
            Total3(k) = Total3(k) + MyDiffEx(Amt - CostAmount, SumItem)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(Amt - CostAmount)
            Total2(k) = Total2(k) + Amt - CostAmount
            Total3(k) = Total3(k) + Amt - CostAmount
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            TotalAmount = TotalAmount + SumItem
            TotalPrice = TotalPrice + Amt
            TotalCost = TotalCost + CostAmount
            
         Wend
         
         k = k + 1
         
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalAmount)
         Total2(k) = Total2(k) + TotalAmount
         Total3(k) = Total3(k) + TotalAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          
          k = k + 1
         
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(MyDiffEx(TotalPrice, TotalAmount))
         Total2(k) = Total2(k) + MyDiffEx(TotalPrice, TotalAmount)
         Total3(k) = Total3(k) + MyDiffEx(TotalPrice, TotalAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalPrice)
         Total2(k) = Total2(k) + TotalPrice
         Total3(k) = Total3(k) + TotalPrice
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
                  
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalAmount)
         Total2(k) = Total2(k) + TotalAmount
         Total3(k) = Total3(k) + TotalAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(MyDiffEx(TotalCost, TotalAmount))
         Total2(k) = Total2(k) + MyDiffEx(TotalCost, TotalAmount)
         Total3(k) = Total3(k) + MyDiffEx(TotalCost, TotalAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalCost)
         Total2(k) = Total2(k) + TotalCost
         Total3(k) = Total3(k) + TotalCost
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalAmount)
         Total2(k) = Total2(k) + TotalAmount
         Total3(k) = Total3(k) + TotalAmount
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(MyDiffEx(TotalPrice - TotalCost, TotalAmount))
         Total2(k) = Total2(k) + MyDiffEx(TotalPrice - TotalCost, TotalAmount)
         Total3(k) = Total3(k) + MyDiffEx(TotalPrice - TotalCost, TotalAmount)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TotalPrice - TotalCost)
         Total2(k) = Total2(k) + TotalPrice - TotalCost
         Total3(k) = Total3(k) + TotalPrice - TotalCost
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
            
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
   Next m_BillingDoc
   
   If (mcolParam("STOCK_TYPE") > 0 Or mcolParam("STOCK_TYPE1") > 0 Or mcolParam("STOCK_TYPE2") > 0) Then
      Call GenerateFooter("      รวม " & PrevKey1, Total2, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   End If
   
   Call GenerateFooter("                        สุทธิ", Total3, m_Details1)
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
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub

