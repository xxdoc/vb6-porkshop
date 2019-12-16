VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling037_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling037_2"
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
Private m_StockCodeColl As Collection
Private m_StockCodeChangeColl As Collection
Private SumPartItemUnitColl As Collection
Private DriverCarTranSportTorColl As Collection
Private DisTinctPartCollection As Collection
Private HeaderStringShow  As String
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
   Set m_StockCodeColl = New Collection
   Set m_StockCodeChangeColl = New Collection
   Set SumPartItemUnitColl = New Collection
   Set DriverCarTranSportTorColl = New Collection
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
   Set m_StockCodeColl = Nothing
   Set m_StockCodeChangeColl = Nothing
   Set SumPartItemUnitColl = Nothing
   Set DriverCarTranSportTorColl = Nothing
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
Private Sub printHeader(PartCollection As Collection)
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
Dim TempBd As CBillingDoc
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Vsp.TextAlign = taLeftMiddle
   Vsp = HeaderStringShow
   Vsp = ""
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "N", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   For Each TempBd In PartCollection
      Call Cf.SetFieldValue(2, "^", TempBd.BILL_DESC, ">", Trim(Str(TempBd.PART_ITEM_ID)))
      Call m_Headers1.AddField(Cf)
   Next TempBd
   
   Call Cf.SetFieldValue(2, "^", "หมายเหตุ", "<")
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
   mblnHeader = False
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
   
   If mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
      Call GetAmountForTranSport(TotalMoves, mcolParam("TO_BILL_DATE"), mcolParam("TO_BILL_DATE"), "(" & PO_DOCTYPE & ")")
   ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
      Call GetAmountForTranSport(TotalMoves, mcolParam("TO_BILL_DATE"), mcolParam("TO_BILL_DATE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")")
   End If
   
   Call LoadStockCode(Nothing, m_StockCodeColl)
   
   Call LoadStockCodeChange(m_StockCodeChangeColl)
   
   Call LoadTranSportByDriverCarTranSportTor(DriverCarTranSportTorColl)
   
   ToDate = "วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & " " & ToDate
                                       
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
Dim ShowUnit As Byte
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_BillingDoc As CBillingDoc
Dim TempBillingDoc As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim TotalPrice As Double


Dim TempBd As CBillingDoc

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
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   
   I = 0
   Set BodyCf = New CReportField
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   
   If mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
      m_BillingDoc.DOCUMENT_TYPE = PO_DOCTYPE
      m_BillingDoc.FROM_DUE_DATE = mcolParam("TO_BILL_DATE")
      m_BillingDoc.TO_DUE_DATE = mcolParam("TO_BILL_DATE")
   ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
      m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")"
      m_BillingDoc.FROM_DATE = mcolParam("TO_BILL_DATE")
      m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
   End If
   Call m_BillingDoc.QueryData(98, Rs, iCount)
   
   I = 0

   While Not Rs.EOF
      I = I + 1
      Call m_Details1.ClearField
      Call m_BillingDoc.PopulateFromRS2(98, Rs)
      
      If PrevKey1 <> Trim(m_BillingDoc.DRIVER_ID & "-" & m_BillingDoc.CAR_LICENSE_ID & "-" & m_BillingDoc.TRANSPORTOR_ID) And I <> 1 Then
'         For ShowUnit = 1 To 10
'            If GenerateFooter("รวม", Total1, m_Details1, ShowUnit) Then
'               Call m_Details1.GetString(1, TempStr1, TempStr2)
'               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'               Vsp.FontBold = True
'               Call Vsp.AddTable(strFormat, "", TempStr2)
'               Vsp.FontBold = False
'            End If
'         Next ShowUnit
'
'         Call GenerateDriverCarTranSportTorDesc(PrevKey1)
         
         I = 1
         
'         HeaderStringShow = "คนขับ  :  " & m_BillingDoc.DRIVER_NAME & "           ทะเบียน  :  " & m_BillingDoc.CAR_LICENSE_NAME & "           ขนส่ง    :  " & m_BillingDoc.TRANSPORTOR_NAME
         
'         Set SumPartItemUnitColl = Nothing
'         Set SumPartItemUnitColl = New Collection
'         Set DisTinctPartCollection = Nothing
'         Set DisTinctPartCollection = New Collection
'         If mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
'            Call GetDistinctForTranSport(DisTinctPartCollection, mcolParam("TO_BILL_DATE"), mcolParam("TO_BILL_DATE"), "(" & PO_DOCTYPE & ")", m_BillingDoc.DRIVER_ID, m_BillingDoc.CAR_LICENSE_ID, m_BillingDoc.TRANSPORTOR_ID)
'         ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
'               Call GetDistinctForTranSport(DisTinctPartCollection, mcolParam("TO_BILL_DATE"), mcolParam("TO_BILL_DATE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")", m_BillingDoc.DRIVER_ID, m_BillingDoc.CAR_LICENSE_ID, m_BillingDoc.TRANSPORTOR_ID)
'         End If
         
'         Vsp.NewPage
         ' รวม
'         For j = 1 To UBound(Total1)
'            Total1(j) = 0
'         Next j
      ElseIf I = 1 Then
         'HeaderStringShow = "คนขับ  :  " & m_BillingDoc.DRIVER_NAME & "           ทะเบียน  :  " & m_BillingDoc.CAR_LICENSE_NAME & "           ขนส่ง    :  " & m_BillingDoc.TRANSPORTOR_NAME
         Set DisTinctPartCollection = Nothing
         Set DisTinctPartCollection = New Collection
         If mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
            Call GetDistinctForTranSport3(DisTinctPartCollection, mcolParam("TO_BILL_DATE"), mcolParam("TO_BILL_DATE"), "(" & PO_DOCTYPE & ")")
         ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
            Call GetDistinctForTranSport3(DisTinctPartCollection, mcolParam("TO_BILL_DATE"), mcolParam("TO_BILL_DATE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")")
         End If
         Call printHeader(DisTinctPartCollection)
         mblnHeader = True
      End If
      
      PrevKey1 = Trim(m_BillingDoc.DRIVER_ID & "-" & m_BillingDoc.CAR_LICENSE_ID & "-" & m_BillingDoc.TRANSPORTOR_ID)
      
      SumRow = 0
      Call m_Details1.ClearField
      
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = I
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = m_BillingDoc.APAR_NAME & "" & AddStringFrontEnd(m_BillingDoc.CUSTOMER_BRANCH_NAME)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      k = 2
      For Each TempBd In DisTinctPartCollection
         Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.DRIVER_ID & "-" & m_BillingDoc.CAR_LICENSE_ID & "-" & m_BillingDoc.TRANSPORTOR_ID & "-" & m_BillingDoc.APAR_MAS_ID & "-" & m_BillingDoc.CUSTOMER_BRANCH & "-" & TempBd.PART_ITEM_ID), True)
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         Amt = MyDiff(TempBillingDoc.TOTAL_AMOUNT, TempBd.UNIT_AMOUNT)
         TempStr = FormatNumberToNull(Amt, 0)
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt
         SumRow = SumRow + Amt
         '------------------------------------------------------------------------------------------------------------------------------------------------
         'หาว่าตอนนี้นั้นมีจำนวนรวมของ แต่ละ รหัสสินค้าและหน่วยอะไรบ้าง
         Dim Scc As CStockCodeChange
         Dim TempSc As CStockCode
         Dim UnitID As Long
         Dim HaveInCustomerFlag As Boolean
         'ค้นหาตาม ลูกค้าและวัตถุดิบ
         
         HaveInCustomerFlag = CheckInCustomer(m_BillingDoc.APAR_MAS_ID, TempBd.PART_ITEM_ID)
         
         For ShowUnit = 1 To 10
            
            If ShowUnit = 1 Then
               UnitID = mcolParam("UNIT_CHANGE1")           'Test กล่อง
            ElseIf ShowUnit = 2 Then
               UnitID = mcolParam("UNIT_CHANGE2")           'Test ตะกร้า
            ElseIf ShowUnit = 3 Then
               UnitID = mcolParam("UNIT_CHANGE3")           'Test ถุง
            ElseIf ShowUnit = 4 Then
               UnitID = mcolParam("UNIT_CHANGE4")           'Test แพ็ค
            ElseIf ShowUnit = 5 Then
               UnitID = mcolParam("UNIT_CHANGE5")           'Test ฟอง
            ElseIf ShowUnit = 6 Then
               UnitID = mcolParam("UNIT_CHANGE6")           '
            ElseIf ShowUnit = 7 Then
               UnitID = mcolParam("UNIT_CHANGE7")           '
            ElseIf ShowUnit = 8 Then
               UnitID = mcolParam("UNIT_CHANGE8")           '
            ElseIf ShowUnit = 9 Then
               UnitID = mcolParam("UNIT_CHANGE9")           '
            ElseIf ShowUnit = 10 Then
               UnitID = mcolParam("UNIT_CHANGE10")           '
            End If
            
            Set Scc = GetObject("CStockCodeChange", m_StockCodeChangeColl, Trim(m_BillingDoc.APAR_MAS_ID & "-" & TempBd.PART_ITEM_ID & "-" & UnitID), False)
            
            If Scc Is Nothing And Not (HaveInCustomerFlag) Then    'ถ้าไม่มีให้ค้นหาและใส่ตามที่ USER เลือก
               Set Scc = GetObject("CStockCodeChange", m_StockCodeChangeColl, Trim(-1 & "-" & TempBd.PART_ITEM_ID & "-" & UnitID), False)
               If Not (Scc Is Nothing) Then
                  '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
                  'โดยเช็คก่อนว่ามีหรือยัง
                  Set TempSc = GetObject("CStockCode", SumPartItemUnitColl, Trim(TempBd.PART_ITEM_ID & "-" & UnitID), False)
                  If TempSc Is Nothing Then
                     Set TempSc = New CStockCode
                     TempSc.STOCK_CODE_ID = TempBd.PART_ITEM_ID
                     TempSc.UNIT_CHANGE_ID = UnitID
                     TempSc.COST_PER_AMOUNT = MyDiff(TempBillingDoc.TOTAL_AMOUNT, TempBd.UNIT_AMOUNT)        'เก็บข้อเป็นหน่วยใหญ่
                     TempSc.UNIT_AMOUNT = Scc.UNIT_CHANGE_AMOUNT           'หน่วยหารจะหน่วยใหญ่เป็นหน่วยที่ใช้ส่ง
                     Call SumPartItemUnitColl.add(TempSc, Trim(TempSc.STOCK_CODE_ID & "-" & UnitID))
                     Set TempSc = Nothing
                  Else
                     TempSc.COST_PER_AMOUNT = TempSc.COST_PER_AMOUNT + MyDiff(TempBillingDoc.TOTAL_AMOUNT, TempBd.UNIT_AMOUNT)
                  End If
                  '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
               End If
            ElseIf Not (Scc Is Nothing) Then  ' ถ้ามีข้อมูลแล้วให้ใส่เข้าได้เลย
               'โดยเช็คก่อนว่ามีหรือยัง
               Set TempSc = GetObject("CStockCode", SumPartItemUnitColl, Trim(TempBd.PART_ITEM_ID & "-" & UnitID), False)
               If TempSc Is Nothing Then
                  Set TempSc = New CStockCode
                  TempSc.STOCK_CODE_ID = TempBd.PART_ITEM_ID
                  TempSc.UNIT_CHANGE_ID = UnitID
                  TempSc.COST_PER_AMOUNT = MyDiff(TempBillingDoc.TOTAL_AMOUNT, TempBd.UNIT_AMOUNT)        'เก็บข้อเป็นหน่วยใหญ่
                  TempSc.UNIT_AMOUNT = Scc.UNIT_CHANGE_AMOUNT           'หน่วยหารจะหน่วยใหญ่เป็นหน่วยที่ใช้ส่ง
                  Call SumPartItemUnitColl.add(TempSc, Trim(TempSc.STOCK_CODE_ID & "-" & UnitID))
                  Set TempSc = Nothing
               Else
                  TempSc.COST_PER_AMOUNT = TempSc.COST_PER_AMOUNT + MyDiff(TempBillingDoc.TOTAL_AMOUNT, TempBd.UNIT_AMOUNT)
               End If
            End If
         Next ShowUnit
         '------------------------------------------------------------------------------------------------------------------------------------------------
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next TempBd
      
      k = k + 1
      Set HeadCf = m_Headers1.Fields(k)
      TempStr = ""
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Call Vsp.AddTable(strFormat, "", TempStr2)
      Rs.MoveNext
   Wend
   For ShowUnit = 1 To 10
      If GenerateFooter("รวม", Total1, m_Details1, ShowUnit) Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
   Next ShowUnit
'   Call GenerateDriverCarTranSportTorDesc(PrevKey1)
   
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
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function

Private Function GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList, k As Byte) As Boolean
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempProfit As Double
Dim TempSale As Double
Dim UnitID  As Long
Dim UnitName As String
Dim StockCode   As CStockCode

      GenerateFooter = False
      
      If k = 1 Then
         UnitID = mcolParam("UNIT_CHANGE1")
         UnitName = mcolParam("UNIT_CHANGE_NAME1")
      ElseIf k = 2 Then
         UnitID = mcolParam("UNIT_CHANGE2")
         UnitName = mcolParam("UNIT_CHANGE_NAME2")
      ElseIf k = 3 Then
         UnitID = mcolParam("UNIT_CHANGE3")
         UnitName = mcolParam("UNIT_CHANGE_NAME3")
      ElseIf k = 4 Then
         UnitID = mcolParam("UNIT_CHANGE4")
         UnitName = mcolParam("UNIT_CHANGE_NAME4")
      ElseIf k = 5 Then
         UnitID = mcolParam("UNIT_CHANGE5")
         UnitName = mcolParam("UNIT_CHANGE_NAME5")
      ElseIf k = 6 Then
         UnitID = mcolParam("UNIT_CHANGE6")
         UnitName = mcolParam("UNIT_CHANGE_NAME6")
      ElseIf k = 7 Then
         UnitID = mcolParam("UNIT_CHANGE7")
         UnitName = mcolParam("UNIT_CHANGE_NAME7")
      ElseIf k = 8 Then
         UnitID = mcolParam("UNIT_CHANGE8")
         UnitName = mcolParam("UNIT_CHANGE_NAME8")
      ElseIf k = 9 Then
         UnitID = mcolParam("UNIT_CHANGE9")
         UnitName = mcolParam("UNIT_CHANGE_NAME9")
      ElseIf k = 10 Then
         UnitID = mcolParam("UNIT_CHANGE10")
         UnitName = mcolParam("UNIT_CHANGE_NAME10")
      End If
      
      If UnitID <= 0 Then
         Exit Function
      End If
      
      Call Details.ClearField
      Set BodyCf = New CReportField
      
      j = 0
      For Each HeadCf In m_Headers1.Fields
         j = j + 1
         If j = 1 Then
            TempStr = ""
         ElseIf j = 2 Then
            TempStr = "รวม " & UnitName
         ElseIf j >= (m_Headers1.Fields.Count) Then
            GenerateFooter = True
            TempStr = ""
         ElseIf j >= 2 Then
            Set StockCode = GetObject("CStockCode", m_StockCodeColl, m_Headers1.Fields(j).TempKey)
            If StockCode.UNIT_ID = UnitID Then
               TempStr = FormatNumberToNull(Tot(j), 0)
            ElseIf StockCode.UNIT_CHANGE_ID = UnitID Then
               TempStr = FormatNumberToNull(Tot(j) * StockCode.UNIT_AMOUNT, 0)
            Else
               Set StockCode = GetObject("CStockCode", SumPartItemUnitColl, m_Headers1.Fields(j).TempKey & "-" & UnitID)
               TempStr = GetMyDiff(StockCode.COST_PER_AMOUNT, StockCode.UNIT_AMOUNT)
            End If
         End If
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      Next HeadCf
      
      Set BodyCf = Nothing
End Function

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
   Vsp.FontSize = iSize
   If mblnHeader Then
      Call printHeader(DisTinctPartCollection)
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
Private Function CheckInCustomer(CustomerID As Long, PartItemID As Long) As Boolean
Dim Scc As CStockCodeChange
Dim I As Long
Dim UnitID As Long
   CheckInCustomer = False
   For I = 1 To 10
      If I = 1 Then
         UnitID = mcolParam("UNIT_CHANGE1")           'Test กล่อง
      ElseIf I = 2 Then
         UnitID = mcolParam("UNIT_CHANGE2")           'Test ตะกร้า
      ElseIf I = 3 Then
         UnitID = mcolParam("UNIT_CHANGE3")           'Test ถุง
      ElseIf I = 4 Then
         UnitID = mcolParam("UNIT_CHANGE4")           'Test แพ็ค
      ElseIf I = 5 Then
         UnitID = mcolParam("UNIT_CHANGE5")           'Test ฟอง
      ElseIf I = 6 Then
         UnitID = mcolParam("UNIT_CHANGE6")           '
      ElseIf I = 7 Then
         UnitID = mcolParam("UNIT_CHANGE7")           '
      ElseIf I = 8 Then
         UnitID = mcolParam("UNIT_CHANGE8")           '
      ElseIf I = 9 Then
         UnitID = mcolParam("UNIT_CHANGE9")           '
      ElseIf I = 10 Then
         UnitID = mcolParam("UNIT_CHANGE10")           '
      End If
      
      Set Scc = GetObject("CStockCodeChange", m_StockCodeChangeColl, Trim(CustomerID & "-" & PartItemID & "-" & UnitID), False)
      If Not (Scc Is Nothing) Then
         CheckInCustomer = True
         Exit Function
      End If
   Next I
End Function
Private Function GetMyDiff(ByVal D1 As Double, ByVal D2 As Double) As String
   If D2 = 0 Or D1 = 0 Then
      GetMyDiff = ""
   ElseIf (D1 Mod D2 > 0) Then
      GetMyDiff = Trim(Str(D1 \ D2)) & "--" & Trim(Str(D1 Mod D2))
   Else
      GetMyDiff = Trim(Str(D1 \ D2))
   End If
End Function
Private Sub GenerateDriverCarTranSportTorDesc(KeyString As String)
Dim TempTran As CTranSportDetail
   
   Set TempTran = GetObject("CTranSportDetail", DriverCarTranSportTorColl, Trim(KeyString), False)
   If Not (TempTran Is Nothing) Then
      Vsp = TempTran.TRANSPORT_DETAIL_DESC
   End If
End Sub
