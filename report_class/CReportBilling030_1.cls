VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling030_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling030_1"
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
Private TempDateCountMonth As Long
Private TempFirstDayOfMonth As Date
Private TempFinalPosition As Long
Private CheckPage As Long
Private TempDateCountTagetDetails As Long
Private TempDateCountNow  As Long
Private TempLastPosition  As Long
Private TempDateCountTaget As Long
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
Private m_Headers3 As CFieldList
Private m_Headers4 As CFieldList
Private m_Headers5 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double

Private DistinctStockNos As Collection
Private FirstDistinctStockNos As Collection
Private OtherMoves As Collection
Private OtherMoveCmps As Collection
Private DistinctStockNos_FB As Collection
Private OtherMoves_FB As Collection
Private OtherMoveCmps_FB As Collection
Private DistinctStockNosConsignment As Collection
Private FirstDistinctStockNosConsignment As Collection
Private DistinctStockNosConsignment_FB As Collection
Private CONSIGNMENT As Collection
Private ConsignmentCmps As Collection
Private CONSIGNMENT_FB As Collection
Private ConsignmentCmps_FB As Collection
Private NotStockNo As Collection
Private m_DistinctMain As Collection
Private m_FirstDistinctMain As Collection
Private m_TempFirstDistinctMain As Collection
Private m_TempFirstDistinctMain2 As Collection
Private m_SaleAmounts As Collection
Private m_SaleAmountExs As Collection
Private m_DistinctMain_FB As Collection
Private m_SaleAmounts_FB As Collection
Private m_SaleAmountExs_FB As Collection
Private m_TagetDetails As Collection
Private m_TagetDetailsSaleCode As Collection
Private m_TagetDetailsSaleCode2 As Collection
Private m_DistinctMain2 As Collection
Private m_FirstDistinctMain2 As Collection
Private m_SaleAmounts2 As Collection
Private m_SaleAmountExs2 As Collection
Private m_DistinctMain_FB2 As Collection
Private m_SaleAmounts_FB2 As Collection
Private m_SaleAmountExs_FB2 As Collection
Private m_TagetDetails2 As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Headers3 = New CFieldList
   Set m_Headers4 = New CFieldList
   Set m_Headers5 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList

   Set OtherMoves = New Collection
   Set OtherMoveCmps = New Collection
   Set DistinctStockNos_FB = New Collection
   Set OtherMoves_FB = New Collection
   Set OtherMoveCmps_FB = New Collection
   Set CONSIGNMENT = New Collection
   Set ConsignmentCmps = New Collection
   Set CONSIGNMENT_FB = New Collection
   Set ConsignmentCmps_FB = New Collection
   Set DistinctStockNos = New Collection
   Set FirstDistinctStockNos = New Collection
   Set DistinctStockNosConsignment = New Collection
   Set FirstDistinctStockNosConsignment = New Collection
   Set DistinctStockNosConsignment_FB = New Collection
   Set NotStockNo = New Collection
   Set m_DistinctMain = New Collection
   Set m_FirstDistinctMain = New Collection
   Set m_TempFirstDistinctMain = New Collection
   Set m_TempFirstDistinctMain2 = New Collection
   Set m_SaleAmounts = New Collection
   Set m_SaleAmountExs = New Collection
   Set m_DistinctMain_FB = New Collection
   Set m_SaleAmounts_FB = New Collection
   Set m_SaleAmountExs_FB = New Collection
   Set m_TagetDetails = New Collection
   Set m_TagetDetailsSaleCode = New Collection
   Set m_TagetDetailsSaleCode2 = New Collection
   Set m_DistinctMain2 = New Collection
   Set m_FirstDistinctMain2 = New Collection
   Set m_SaleAmounts2 = New Collection
   Set m_SaleAmountExs2 = New Collection
   Set m_DistinctMain_FB2 = New Collection
   Set m_SaleAmounts_FB2 = New Collection
   Set m_SaleAmountExs_FB2 = New Collection
   Set m_TagetDetails2 = New Collection
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
   Set m_Headers3 = Nothing
   Set m_Headers4 = Nothing
   Set m_Headers5 = Nothing
   Set m_Details2 = Nothing

   Set OtherMoves = Nothing
   Set OtherMoveCmps = Nothing
   Set DistinctStockNos_FB = Nothing
   Set OtherMoves_FB = Nothing
   Set OtherMoveCmps_FB = Nothing
   Set CONSIGNMENT = Nothing
   Set ConsignmentCmps = Nothing
   Set CONSIGNMENT_FB = Nothing
   Set ConsignmentCmps_FB = Nothing
   Set DistinctStockNos = Nothing
   Set FirstDistinctStockNos = Nothing
   Set DistinctStockNosConsignment = Nothing
   Set FirstDistinctStockNosConsignment = Nothing
   Set DistinctStockNosConsignment_FB = Nothing
   Set NotStockNo = Nothing
   Set m_DistinctMain = Nothing
   Set m_FirstDistinctMain = Nothing
   Set m_TempFirstDistinctMain = Nothing
   Set m_TempFirstDistinctMain2 = Nothing
   Set m_SaleAmounts = Nothing
   Set m_SaleAmountExs = Nothing
   Set m_DistinctMain_FB = Nothing
   Set m_SaleAmounts_FB = Nothing
   Set m_SaleAmountExs_FB = Nothing
   Set m_TagetDetails = Nothing
   Set m_TagetDetailsSaleCode = Nothing
   Set m_TagetDetailsSaleCode2 = Nothing
   Set m_DistinctMain2 = Nothing
   Set m_FirstDistinctMain2 = Nothing
   Set m_SaleAmounts2 = Nothing
   Set m_SaleAmountExs2 = Nothing
   Set m_DistinctMain_FB2 = Nothing
   Set m_SaleAmounts_FB2 = Nothing
   Set m_SaleAmountExs_FB2 = Nothing
   Set m_TagetDetails2 = Nothing
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
   
   Call Cf.SetFieldValue(4, "^", "รายการเอียด", "<")
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
   
   If Month(mcolParam("FROM_BILL_DATE")) = Month(mcolParam("TO_BILL_DATE")) Then
      Call Cf.SetFieldValue(4, "^", "ต้นทุน " & Format(Day(mcolParam("FROM_BILL_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   ElseIf Year(mcolParam("FROM_BILL_DATE")) = Year(mcolParam("TO_BILL_DATE")) Then
      Call Cf.SetFieldValue(4, "^", "ต้นทุน " & Format(Day(mcolParam("FROM_BILL_DATE")), "00") & "/" & Format(Month(mcolParam("FROM_BILL_DATE")), "00") & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   Else
      Call Cf.SetFieldValue(4, "^", "ต้นทุน " & DateToStringExtEx2(mcolParam("FROM_BILL_DATE")) & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   End If
   Call m_Headers0.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "กำไร", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "รายการสินค้า", "<")
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

   Call Cf.SetFieldValue(0, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">", , True)
   Call m_Headers1.AddField(Cf)
   
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
Private Sub printHeader2()
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
Dim TempDate As Long

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers2.ClearField
   Call m_Headers3.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(6, "^", "ประเภทสินค้า", "<")
   Call m_Headers2.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "", "<")
   Call m_Headers3.AddField(Cf)
   
   iCount = 0
   TempDate = 1
   TempFinalPosition = 0
   For j = 1 To TempDateCountMonth
      If Weekday(TempFirstDayOfMonth) = 1 Or j = TempDateCountMonth Then
         iCount = iCount + 1
         TempFinalPosition = TempFinalPosition + 2
         Call Cf.SetFieldValue(4, "^", "Week # " & iCount & " (" & TempDate & "-" & j & ")", ">")
         Call m_Headers2.AddField(Cf)
         
         Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
         Call m_Headers3.AddField(Cf)
         
         Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
         Call m_Headers3.AddField(Cf)
         TempDate = j + 1
      End If
      TempFirstDayOfMonth = DateAdd("d", 1, TempFirstDayOfMonth)
   Next j
   
   TempFinalPosition = TempFinalPosition + 2
   Call Cf.SetFieldValue(4, "^", "รวม", ">")
   Call m_Headers2.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
   Call m_Headers3.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "จำนวนเงิน", ">")
   Call m_Headers3.AddField(Cf)
   
   
   Set Cf = Nothing
               
   Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers3.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader3()
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
Dim DateCount  As Long
Dim TempDateString As String
Dim TempFromDate  As Date
Dim TempToDate   As Date
Dim FD As Date
Dim LD As Date
Dim D As String

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers4.ClearField
   Call m_Headers5.ClearField
   
   Set Cf = New CReportField
   
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = DateAdd("D", DateCount - 1, mcolParam("FROM_DATE"))
   Call GetFirstLastDate(TempFromDate, FD, LD)
   TempDateCountTagetDetails = DateDiff("D", FD, LD) + 1 - Val(mcolParam("HOLIDAY"))
   TempDateCountNow = DateCount
   
   Call Cf.SetFieldValue(5, "^", "รายละเอียด", "<")
   Call m_Headers4.AddField(Cf)
    '------------------------------------------------------------------------------------------>          เป้าการขายเดือน
   Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
   Call m_Headers4.AddField(Cf)
   
   If TempDateCountNow <> TempDateCountTagetDetails Then
      Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
      Call m_Headers4.AddField(Cf)
   End If

   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))  'ตัดวงเล็บออก
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   '------------------------------------------------------------------------------------------>          ยอดขาย
    While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      Call Cf.SetFieldValue(2, "^", "Actual", ">")
      Call m_Headers4.AddField(Cf)
   Wend
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers4.AddField(Cf)
   '------------------------------------------------------------------------------------------>
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   
   Call Cf.SetFieldValue(5, "^", "", "<")
   Call m_Headers5.AddField(Cf)
   '------------------------------------------------------------------------------------------>          เป้าการขายเดือน
   Call Cf.SetFieldValue(2, "^", "เป้าเดือน " & IntToThaiMonth(Format(mcolParam("MONTH_ID"), "0"), 1), "<")
   Call m_Headers5.AddField(Cf)
   
    If TempDateCountNow <> TempDateCountTagetDetails Then
      Call Cf.SetFieldValue(2, "^", "เป้า " & TempDateCountNow & " วัน", "<")
      Call m_Headers5.AddField(Cf)
   End If
   '------------------------------------------------------------------------------------------>          ยอดขาย
   TempFromDate = mcolParam("FROM_DATE")
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
      TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)
      
      Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
      Call m_Headers5.AddField(Cf)
         
      TempFromDate = DateAdd("D", 1, TempToDate)
   Wend

   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers5.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ผลต่าง", ">")
   Call m_Headers5.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers5.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers4.AddField(Cf)
   '------------------------------------------------------------------------------------------>
   Set Cf = Nothing
   
   Call m_Headers5.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers4.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader4()
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
Dim DateCount  As Long
Dim TempDateString As String
Dim TempFromDate  As Date
Dim TempToDate   As Date
Dim FD As Date
Dim LD As Date
Dim D As String

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers4.ClearField
   Call m_Headers5.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(4, "^", "เป้าการขายเดือน " & Month(mcolParam("FROM_DATE")), "<")
   Call m_Headers5.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "รหัสพนักงานขาย / พนักงานขาย", "<")
   Call m_Headers4.AddField(Cf)

   Call Cf.SetFieldValue(8, "^", "วันที่ " & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   Call m_Headers5.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เป้าการขาย", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Diff", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers4.AddField(Cf)
   'สะสม
   Call GetFirstLastDate(mcolParam("TO_BILL_DATE"), FD, LD)
   Call Cf.SetFieldValue(8, "^", "สะสม " & Day(FD) & "-" & DateToStringExtEx2(mcolParam("TO_BILL_DATE")), ">")
   Call m_Headers5.AddField(Cf)
   
   DateCount = DateDiff("D", FD, mcolParam("TO_BILL_DATE")) + 1 - Val(mcolParam("HOLIDAY2"))
   Call Cf.SetFieldValue(2, "^", "เป้าการขาย " & DateCount & " วัน", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "Diff", ">")
   Call m_Headers4.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "%", ">")
   Call m_Headers4.AddField(Cf)
   Set Cf = Nothing
   
   Call m_Headers5.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers4.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader5()
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

   Call m_Headers5.ClearField
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "", "^")
   Call m_Headers5.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดสะสมรายวัน", ">")
   Call m_Headers5.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดสะสมรวม", ">")
   Call m_Headers5.AddField(Cf)
   
   Set Cf = Nothing
   
   Call m_Headers5.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid - 8000, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Set TempRs = Nothing
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub
Private Sub printHeader6()
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
Dim DateCount  As Long
Dim TempDateString As String
Dim TempFromDate  As Date
Dim TempToDate   As Date
Dim FD As Date
Dim LD As Date
Dim D As String

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

   Call m_Headers5.ClearField
   
   Set Cf = New CReportField
   
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = DateAdd("D", DateCount - 1, mcolParam("FROM_DATE"))
   
   Call Cf.SetFieldValue(2, "^", "", "<")
   Call m_Headers5.AddField(Cf)
   '------------------------------------------------------------------------------------------>          ยอดขาย
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))  'ตัดวงเล็บออก
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   TempFromDate = mcolParam("FROM_DATE")
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
      TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)
      
      Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
      Call m_Headers5.AddField(Cf)
         
      TempFromDate = DateAdd("D", 1, TempToDate)
   Wend

   Call Cf.SetFieldValue(2, "^", "รวม", ">")
   Call m_Headers5.AddField(Cf)
   '------------------------------------------------------------------------------------------>
   Set Cf = Nothing
   
   Call m_Headers5.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid - 4000, alngX)
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
Dim FromDate As Date
Dim ToDate As Date
Dim FD As Date
Dim LD As Date
Dim D As String
Dim DateCount As Long
   
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
   
   Call GetDistinctStockCode(FirstDistinctStockNos, FromDate, ToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", 1, , 1)
   Call GetSaleAmountStockCodeDocType(OtherMoves, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", 1, , 1)
   Call GetSaleAmountStockCodeDocType(OtherMoveCmps, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", 1, , 1)
   'F&B
   Call GetDistinctStockCodeForFAndB(DistinctStockNos_FB, FromDate, ToDate, mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", 1)
   Call GetSaleAmountStockCodeDocTypeForFAndB(OtherMoves_FB, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", 1)
   Call GetSaleAmountStockCodeDocTypeForFAndB(OtherMoveCmps_FB, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", 1)
   Call SumDistinctStockNos(FirstDistinctStockNos, DistinctStockNos_FB, DistinctStockNos)
   
   If mcolParam("CONSIGNMENT") <> 0 Then
      'ฝากขาย
      Call GetDistinctStockCode(FirstDistinctStockNosConsignment, FromDate, ToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & PO_DOCTYPE & ")", 1, 1, 1)
      Call GetSaleAmountStockCodeDocType(CONSIGNMENT, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & PO_DOCTYPE & ")", 1, 1, 1)
      Call GetSaleAmountStockCodeDocType(ConsignmentCmps, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & PO_DOCTYPE & ")", 1, 1, 1)
      'F&B
      Call GetDistinctStockCodeForFAndB(DistinctStockNosConsignment_FB, FromDate, ToDate, mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), "(" & PO_DOCTYPE & ")", 1, 1)
      Call GetSaleAmountStockCodeDocTypeForFAndB(CONSIGNMENT_FB, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), "(" & PO_DOCTYPE & ")", 1, 1)
      Call GetSaleAmountStockCodeDocTypeForFAndB(ConsignmentCmps_FB, mcolParam("FROM_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), "(" & PO_DOCTYPE & ")", 1, 1)
      Call SumDistinctStockNos(FirstDistinctStockNosConsignment, DistinctStockNosConsignment_FB, DistinctStockNosConsignment)
   End If
   'สำหรับรายละเอียดจุลินทรีย์
   Call GetFirstLastDate(mcolParam("FROM_BILL_DATE"), FromDate, ToDate)
   TempDateCountMonth = DateDiff("D", FromDate, ToDate) + 1
   TempFirstDayOfMonth = FromDate
   'ส่วนของรายงานที่ 2
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
   FromDate = mcolParam("FROM_DATE")
   ToDate = DateAdd("D", DateCount - 1, mcolParam("FROM_DATE"))
   Call GetFirstLastDate(FromDate, FD, LD)
   TempDateCountTagetDetails = DateDiff("D", FD, LD) + 1 - Val(mcolParam("HOLIDAY"))
   TempDateCountNow = DateCount
   
   Call LoadTagetDetailEmpStock(m_TagetDetails, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")))
   Call LoadTagetDetailSaleCode(m_TagetDetailsSaleCode, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")))
   
   Call GetDistinctEmpStockCodeDocTypeFree(m_TempFirstDistinctMain, FromDate, ToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), 1)
   Call GetSaleAmountEmpStockCodeDocTypeFreeDate(m_SaleAmounts, FromDate, ToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), 1, 1)
   Call GetSaleAmountEmpStockCodeDocTypeFreeExDate(m_SaleAmountExs, FromDate, ToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), 1, 1)
   'F&B
   Call GetDistinctEmpStockCodeDocTypeFreeForFAndB(m_DistinctMain_FB, FromDate, ToDate, mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_SALE_CODE2"), mcolParam("TO_SALE_CODE2"))
   Call GetSaleAmountEmpStockCodeDocTypeFreeDateForFAndB(m_SaleAmounts_FB, FromDate, ToDate, mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), mcolParam("FROM_SALE_CODE2"), mcolParam("TO_SALE_CODE2"), 1)
   Call GetSaleAmountEmpStockCodeDocTypeFreeExDateForFAndB(m_SaleAmountExs_FB, FromDate, ToDate, mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"), mcolParam("FROM_APAR_CODE2"), mcolParam("TO_APAR_CODE2"), mcolParam("FROM_SALE_CODE2"), mcolParam("TO_SALE_CODE2"), 1)
   
   Call SumTempDistinctMain(m_TempFirstDistinctMain, m_TagetDetailsSaleCode, m_FirstDistinctMain)
   Call SumDistinctMain(m_FirstDistinctMain, m_DistinctMain_FB, m_DistinctMain)
   'ส่วนที่3
   Call GetFirstLastDate(mcolParam("TO_BILL_DATE"), FromDate, ToDate)
   TempDateCountTaget = DateDiff("D", FromDate, ToDate) + 1 - Val(mcolParam("HOLIDAY"))
   
   Call LoadTagetDetailEmp(m_TagetDetails2, Trim((Val(Year(mcolParam("TO_BILL_DATE")))) & Format(Month(mcolParam("TO_BILL_DATE")), "00")))
   Call LoadTagetDetailSaleCode2(m_TagetDetailsSaleCode2, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")))
   
   Call GetDistinctEmpDocTypeFree(m_TempFirstDistinctMain2, FromDate, mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), 1)
   Call GetSaleAmountEmp(m_SaleAmounts2, FromDate, mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), 1)
   Call GetSaleAmountEmpDateFree(m_SaleAmountExs2, FromDate, mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), 1)
   'F&B
   Call GetDistinctEmpDocTypeFreeForFAndB(m_DistinctMain_FB2, FromDate, mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE2"), mcolParam("TO_SALE_CODE2"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"))
   Call GetSaleAmountEmpForFAndB(m_SaleAmounts_FB2, FromDate, mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE2"), mcolParam("TO_SALE_CODE2"), mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"))
   Call GetSaleAmountEmpDateFreeForFAndB(m_SaleAmountExs_FB2, FromDate, mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE2"), mcolParam("TO_SALE_CODE2"), , mcolParam("FROM_STOCK_NO2"), mcolParam("TO_STOCK_NO2"))
   
   Call SumTempDistinctMain2(m_TempFirstDistinctMain2, m_TagetDetailsSaleCode2, m_FirstDistinctMain2)
   Call SumDistinctMain2(m_FirstDistinctMain2, m_DistinctMain_FB2, m_DistinctMain2)

   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       "จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A")
                                       
                                       
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
Dim Amt2 As Double
Dim m_BillingDoc As CBillingDoc
Dim TempBillingDoc As CBillingDoc
Dim TempBillingDoc2 As CBillingDoc
Dim TempBillingDoc3 As CBillingDoc
Dim TempBillingDocFB As CBillingDoc
Dim TempBillingDocFB2 As CBillingDoc
Dim TempBillingDocFB3 As CBillingDoc
Dim TempBd As CBillingDoc
Dim Tg As CTagetDetail
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim Total2(25) As Double
Dim Total3(25) As Double
Dim Total4(25) As Double
Dim Total5(25) As Double
Dim Total6(25) As Double
Dim Total7(25) As Double
Dim SumItem As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As Long
Dim TotalPrice As Double
Dim TotalPrice2 As Double
Dim CostAmount As Double
Dim FromDate As Date
Dim ToDate As Date
Dim TempDate As Date
Dim CountGroup As Long
Dim CountColor As Long
Dim DateCount As Long
Dim TempFirstPosition As Long
Dim TempDateString As String
Dim CountDate As Long
Dim TempDateCountRepost3  As Long

Dim k As Long
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
      Total4(j) = 0
      Total5(j) = 0
      Total6(j) = 0
      Total7(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   I = 0
   Set BodyCf = New CReportField

   I = 0
   k = 0
   CountGroup = 0
   CountColor = 0
   
   For Each m_BillingDoc In DistinctStockNos
      k = k + 1
      If (PrevKey2 <> m_BillingDoc.STOCK_GROUP_NAME) And (I > 0) Then
         
         If PrevKey2 <> "จุลินทรีย์" Then
            Call GenerateFooter(PrevKey2, Total4, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            CountColor = CountColor + 1               'เปลี่ยนสี
            If CountColor > 2 Then
               Vsp.TextColor = RGB(0, 150, 0)
               If CountColor > 3 Then
                  CountColor = 0
               End If
            End If
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            Vsp.TextColor = RGB(0, 0, 0)
         End If
         
         'แสดงกลุ่มทั้ง 3 กลุ่ม
         If m_BillingDoc.STOCK_GROUP_CODE = "01" Or m_BillingDoc.STOCK_GROUP_CODE = "02" Or m_BillingDoc.STOCK_GROUP_CODE = "03" Then
               CountGroup = 1
         ElseIf m_BillingDoc.STOCK_GROUP_CODE = "04" Or m_BillingDoc.STOCK_GROUP_CODE = "05" Then
               CountGroup = 2
         ElseIf m_BillingDoc.STOCK_GROUP_CODE = "06" Or m_BillingDoc.STOCK_GROUP_CODE = "07" Or m_BillingDoc.STOCK_GROUP_CODE = "08" Or m_BillingDoc.STOCK_GROUP_CODE = "09" Then
               CountGroup = 3
        ElseIf m_BillingDoc.STOCK_GROUP_CODE = "10" Then
               CountGroup = 4
         End If
            
         If PrevKey4 <> CountGroup Then
            Call GenerateFooterEnd("รวมกลุ่ม " & PrevKey4, Total5, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.TextColor = RGB(0, 0, 255)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.TextColor = RGB(0, 0, 0)
            Vsp.FontBold = False
               
            If CountGroup = 1 Then
               Vsp = "กลุ่ม :  1"
            ElseIf CountGroup = 2 Then
               Vsp = "กลุ่ม :  2"
            ElseIf CountGroup = 3 Then
               Vsp = "กลุ่ม :  3"
            ElseIf CountGroup = 4 Then
               Vsp = "กลุ่ม :  4"
            End If
            PrevKey4 = CountGroup
               
            For j = 1 To UBound(Total5)
               Total5(j) = 0
            Next j
         End If
             
         For j = 1 To UBound(Total1)
            Total2(j) = 0
            Total4(j) = 0
         Next j
      ElseIf (PrevKey1 <> m_BillingDoc.STOCK_TYPE_NAME) And (I > 0) Then
            For j = 1 To UBound(Total1)
               Total2(j) = 0
            Next j
      ElseIf I = 0 Then
         If m_BillingDoc.STOCK_GROUP_CODE = "01" Or m_BillingDoc.STOCK_GROUP_CODE = "02" Or m_BillingDoc.STOCK_GROUP_CODE = "03" Then
            Vsp = "กลุ่ม :  1"
            PrevKey4 = 1
         ElseIf m_BillingDoc.STOCK_GROUP_CODE = "04" Or m_BillingDoc.STOCK_GROUP_CODE = "05" Then
            Vsp = "กลุ่ม :  2"
            PrevKey4 = 2
         ElseIf m_BillingDoc.STOCK_GROUP_CODE = "06" Or m_BillingDoc.STOCK_GROUP_CODE = "07" Or m_BillingDoc.STOCK_GROUP_CODE = "08" Or m_BillingDoc.STOCK_GROUP_CODE = "09" Then
            Vsp = "กลุ่ม :  3"
            PrevKey4 = 3
         ElseIf m_BillingDoc.STOCK_GROUP_CODE = "10" Then
            Vsp = "กลุ่ม :  4"
            PrevKey4 = 4
         End If
      End If
  
      I = I + 1
      Call m_Details1.ClearField
      Vsp.TextColor = RGB(0, 0, 0)
         
      PrevKey1 = m_BillingDoc.STOCK_TYPE_NAME
      PrevKey2 = m_BillingDoc.STOCK_GROUP_NAME

      Set TempBillingDoc = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE), True)
      Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE), True)
      Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoveCmps, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE), True)
         
      SumItem = TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
      
      'F&B
      Set TempBillingDocFB = GetObject("CBillingDoc", OtherMoveCmps_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE), True)
      Set TempBillingDocFB2 = GetObject("CBillingDoc", OtherMoveCmps_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE), True)
      Set TempBillingDocFB3 = GetObject("CBillingDoc", OtherMoveCmps_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE), True)
         
      SumItem = SumItem + (TempBillingDocFB.TOTAL_AMOUNT + TempBillingDocFB2.TOTAL_AMOUNT - TempBillingDocFB3.TOTAL_AMOUNT)
      
      Total2(2) = Total2(2) + SumItem
      Total3(2) = Total3(2) + SumItem
      Total4(2) = Total4(2) + SumItem
      Total5(2) = Total5(2) + SumItem
      Total6(2) = Total6(2) + SumItem
          
      Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
      Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
      Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
      'F&B
      Amt = Amt + (TempBillingDocFB.TOTAL_PRICE - TempBillingDocFB.DISCOUNT_AMOUNT - TempBillingDocFB.EXT_DISCOUNT_AMOUNT)
      Amt = Amt + TempBillingDocFB2.TOTAL_PRICE - TempBillingDocFB2.DISCOUNT_AMOUNT - TempBillingDocFB2.EXT_DISCOUNT_AMOUNT
      Amt = Amt - (TempBillingDocFB3.TOTAL_PRICE - TempBillingDocFB3.DISCOUNT_AMOUNT - TempBillingDocFB3.EXT_DISCOUNT_AMOUNT)
      
      Total2(3) = Total2(3) + MyDiffEx(Amt, SumItem)
      Total3(3) = Total3(3) + MyDiffEx(Amt, SumItem)
      Total4(3) = Total4(3) + MyDiffEx(Amt, SumItem)
      Total5(3) = Total5(3) + MyDiffEx(Amt, SumItem)
      Total6(3) = Total6(3) + MyDiffEx(Amt, SumItem)

      Total2(4) = Total2(4) + Amt
      Total3(4) = Total3(4) + Amt
      Total4(4) = Total4(4) + Amt
      Total5(4) = Total5(4) + Amt
      Total6(4) = Total5(4) + Amt

      Set TempBillingDoc = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE), True)
      Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE), True)
      Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE), True)

      SumItem = TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
      'F&B
      Set TempBillingDocFB = GetObject("CBillingDoc", OtherMoves_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE), True)
      Set TempBillingDocFB2 = GetObject("CBillingDoc", OtherMoves_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE), True)
      Set TempBillingDocFB3 = GetObject("CBillingDoc", OtherMoves_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE), True)
         
      SumItem = SumItem + (TempBillingDocFB.TOTAL_AMOUNT + TempBillingDocFB2.TOTAL_AMOUNT - TempBillingDocFB3.TOTAL_AMOUNT)
      
      Total2(5) = Total2(5) + SumItem
      Total3(5) = Total3(5) + SumItem
      Total4(5) = Total4(5) + SumItem
      Total5(5) = Total5(5) + SumItem
      Total6(5) = Total6(5) + SumItem
          
      Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
      Amt = Amt + TempBillingDoc2.TOTAL_PRICE - TempBillingDoc2.DISCOUNT_AMOUNT - TempBillingDoc2.EXT_DISCOUNT_AMOUNT
      Amt = Amt - (TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT)
      'F&B
      Amt = Amt + TempBillingDocFB.TOTAL_PRICE - TempBillingDocFB.DISCOUNT_AMOUNT - TempBillingDocFB.EXT_DISCOUNT_AMOUNT
      Amt = Amt + TempBillingDocFB2.TOTAL_PRICE - TempBillingDocFB2.DISCOUNT_AMOUNT - TempBillingDocFB2.EXT_DISCOUNT_AMOUNT
      Amt = Amt - (TempBillingDocFB3.TOTAL_PRICE - TempBillingDocFB3.DISCOUNT_AMOUNT - TempBillingDocFB3.EXT_DISCOUNT_AMOUNT)
      
      If TempBillingDoc3.TOTAL_PRICE - TempBillingDoc3.DISCOUNT_AMOUNT - TempBillingDoc3.EXT_DISCOUNT_AMOUNT > 0 Then
         Debug.Print
      End If
      If TempBillingDocFB3.TOTAL_PRICE - TempBillingDocFB3.DISCOUNT_AMOUNT - TempBillingDocFB3.EXT_DISCOUNT_AMOUNT > 0 Then
         Debug.Print
      End If
      
      Total2(6) = Total2(6) + MyDiffEx(Amt, SumItem)
      Total3(6) = Total3(6) + MyDiffEx(Amt, SumItem)
      Total4(6) = Total4(6) + MyDiffEx(Amt, SumItem)
      Total5(6) = Total5(6) + MyDiffEx(Amt, SumItem)
      Total6(6) = Total6(6) + MyDiffEx(Amt, SumItem)
   
      Total2(7) = Total2(7) + Amt
      Total3(7) = Total3(7) + Amt
      Total4(7) = Total4(7) + Amt
      Total5(7) = Total5(7) + Amt
      Total6(7) = Total6(7) + Amt

      Total2(8) = Total2(8) + SumItem
      Total3(8) = Total3(8) + SumItem
      Total4(8) = Total4(8) + SumItem
      Total5(8) = Total5(8) + SumItem
      Total6(8) = Total6(8) + SumItem
         
      CostAmount = (TempBillingDoc.TOTAL_INCLUDE_PRICE) + (TempBillingDoc2.TOTAL_INCLUDE_PRICE) - (TempBillingDoc3.TOTAL_INCLUDE_PRICE)
      'F&B
      CostAmount = CostAmount + (TempBillingDocFB.TOTAL_INCLUDE_PRICE) + (TempBillingDocFB2.TOTAL_INCLUDE_PRICE) - (TempBillingDocFB3.TOTAL_INCLUDE_PRICE)

      Total2(9) = Total2(9) + MyDiffEx(CostAmount, SumItem)
      Total3(9) = Total3(9) + MyDiffEx(CostAmount, SumItem)
      Total4(9) = Total4(9) + MyDiffEx(CostAmount, SumItem)
      Total5(9) = Total5(9) + MyDiffEx(CostAmount, SumItem)
      Total6(9) = Total6(9) + MyDiffEx(CostAmount, SumItem)

      Total2(10) = Total2(10) + CostAmount
      Total3(10) = Total3(10) + CostAmount
      Total4(10) = Total4(10) + CostAmount
      Total5(10) = Total5(10) + CostAmount
      Total6(10) = Total6(10) + CostAmount

      Total2(11) = Total2(11) + SumItem
      Total3(11) = Total3(11) + SumItem
      Total4(11) = Total4(11) + SumItem
      Total5(11) = Total5(11) + SumItem
      Total6(11) = Total6(11) + SumItem

      Total2(12) = Total2(12) + MyDiffEx(Amt - CostAmount, SumItem)
      Total3(12) = Total3(12) + MyDiffEx(Amt - CostAmount, SumItem)
      Total4(12) = Total4(12) + MyDiffEx(Amt - CostAmount, SumItem)
      Total5(12) = Total5(12) + MyDiffEx(Amt - CostAmount, SumItem)
      Total6(12) = Total6(12) + MyDiffEx(Amt - CostAmount, SumItem)

      Total2(13) = Total2(13) + Amt - CostAmount
      Total3(13) = Total3(13) + Amt - CostAmount
      Total4(13) = Total4(13) + Amt - CostAmount
      Total5(13) = Total5(13) + Amt - CostAmount
      Total6(13) = Total6(13) + Amt - CostAmount
      
      'แสดงยอดจุลินทรีย์  Requirement พี่นิ่ม 11/10/56
      If CountGroup = 2 And m_BillingDoc.STOCK_GROUP_CODE = "04" Then
            Call GenerateFooter("[" & m_BillingDoc.STOCK_NO & "] " & m_BillingDoc.STOCK_DESC, Total6, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
'            Vsp.TextColor = RGB(0, 0, 0)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
      End If
      For j = 1 To UBound(Total1)
         Total6(j) = 0
      Next j
      
   Next m_BillingDoc

   If PrevKey2 <> "จุลินทรีย์" Then
      Call GenerateFooter(PrevKey2, Total4, m_Details1)
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
   
   Call GenerateFooterEnd("รวมกลุ่ม " & PrevKey4, Total5, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.TextColor = RGB(0, 0, 255)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.TextColor = RGB(0, 0, 0)
   Vsp.FontBold = False
   
   Call GenerateFooterEnd("                        สุทธิ", Total3, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TextColor = RGB(0, 0, 255)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
    If mcolParam("CONSIGNMENT") <> 0 Then
      'ฝากขาย
      Vsp = ""
      Vsp = "ฝากขาย"
      Vsp.TextColor = RGB(0, 0, 0)
      I = 0
      k = 0
      CountColor = 0
      For j = 1 To UBound(Total1)
         Total1(j) = 0
         Total2(j) = 0
         Total3(j) = 0
         Total4(j) = 0
      Next j
      
      For Each m_BillingDoc In DistinctStockNosConsignment                    ' กลุ่มฝากขาย
         k = k + 1
         If (PrevKey2 <> m_BillingDoc.STOCK_GROUP_NAME) And (I > 0) Then
            Call GenerateFooter(PrevKey2, Total4, m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            CountColor = CountColor + 1
            If CountColor > 2 Then
               Vsp.TextColor = RGB(0, 150, 0)
               If CountColor > 3 Then
                  CountColor = 0
               End If
            End If
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            Vsp.TextColor = RGB(0, 0, 0)
                     
            For j = 1 To UBound(Total1)
               Total2(j) = 0
               Total4(j) = 0
            Next j
         ElseIf (PrevKey1 <> m_BillingDoc.STOCK_TYPE_NAME) And (I > 0) Then
            For j = 1 To UBound(Total1)
               Total2(j) = 0
            Next j
         End If
                      
         I = I + 1
         Call m_Details1.ClearField
         Vsp.TextColor = RGB(0, 0, 0)
                  
         PrevKey1 = m_BillingDoc.STOCK_TYPE_NAME
         PrevKey2 = m_BillingDoc.STOCK_GROUP_NAME
         Debug.Print m_BillingDoc.APAR_CODE
   
         Set TempBillingDoc = GetObject("CBillingDoc", ConsignmentCmps, Trim(m_BillingDoc.STOCK_NO & "-" & PO_DOCTYPE), True)
         SumItem = TempBillingDoc.TOTAL_AMOUNT
         'F&B
         Set TempBillingDocFB = GetObject("CBillingDoc", ConsignmentCmps_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & PO_DOCTYPE), True)
         If TempBillingDocFB.TOTAL_AMOUNT <> 0 Then
            SumItem = SumItem + TempBillingDocFB.TOTAL_AMOUNT
         End If
   
         Total2(2) = Total2(2) + SumItem
         Total3(2) = Total3(2) + SumItem
         Total4(2) = Total4(2) + SumItem
                   
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         'F&B
         If TempBillingDocFB.TOTAL_PRICE <> 0 Then
            Amt = Amt + TempBillingDocFB.TOTAL_PRICE - TempBillingDocFB.DISCOUNT_AMOUNT - TempBillingDocFB.EXT_DISCOUNT_AMOUNT
         End If
   
         Total2(3) = Total2(3) + MyDiffEx(Amt, SumItem)
         Total3(3) = Total3(3) + MyDiffEx(Amt, SumItem)
         Total4(3) = Total4(3) + MyDiffEx(Amt, SumItem)
   
         Total2(4) = Total2(4) + Amt
         Total3(4) = Total3(4) + Amt
         Total4(4) = Total4(4) + Amt
                  
         Set TempBillingDoc = GetObject("CBillingDoc", CONSIGNMENT, Trim(m_BillingDoc.STOCK_NO & "-" & PO_DOCTYPE), True)
         SumItem = TempBillingDoc.TOTAL_AMOUNT
         'F&B
         Set TempBillingDocFB = GetObject("CBillingDoc", CONSIGNMENT_FB, Trim(m_BillingDoc.JOINT_CODE & "-" & PO_DOCTYPE), True)
         If TempBillingDocFB.TOTAL_AMOUNT <> 0 Then
            SumItem = SumItem + TempBillingDocFB.TOTAL_AMOUNT
         End If
   
         Total2(5) = Total2(5) + SumItem
         Total3(5) = Total3(5) + SumItem
         Total4(5) = Total4(5) + SumItem
   
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         If TempBillingDocFB.TOTAL_PRICE <> 0 Then
            Amt = Amt + TempBillingDocFB.TOTAL_PRICE - TempBillingDocFB.DISCOUNT_AMOUNT - TempBillingDocFB.EXT_DISCOUNT_AMOUNT
         End If
         
         Total2(6) = Total2(6) + MyDiffEx(Amt, SumItem)
         Total3(6) = Total3(6) + MyDiffEx(Amt, SumItem)
         Total4(6) = Total4(6) + MyDiffEx(Amt, SumItem)
   
         Total2(7) = Total2(7) + Amt
         Total3(7) = Total3(7) + Amt
         Total4(7) = Total4(7) + Amt
   
         Total2(8) = Total2(8) + SumItem
         Total3(8) = Total3(8) + SumItem
         Total4(8) = Total4(8) + SumItem
                  
         CostAmount = (TempBillingDoc.TOTAL_INCLUDE_PRICE) + TempBillingDocFB.TOTAL_INCLUDE_PRICE
   
         Total2(9) = Total2(9) + MyDiffEx(CostAmount, SumItem)
         Total3(9) = Total3(9) + MyDiffEx(CostAmount, SumItem)
         Total4(9) = Total4(9) + MyDiffEx(CostAmount, SumItem)
   
         Total2(10) = Total2(10) + CostAmount
         Total3(10) = Total3(10) + CostAmount
         Total4(10) = Total4(10) + CostAmount
               
         Total2(11) = Total2(11) + SumItem
         Total3(11) = Total3(11) + SumItem
         Total4(11) = Total4(11) + SumItem
   
         Total2(12) = Total2(12) + MyDiffEx(Amt - CostAmount, SumItem)
         Total3(12) = Total3(12) + MyDiffEx(Amt - CostAmount, SumItem)
         Total4(12) = Total4(12) + MyDiffEx(Amt - CostAmount, SumItem)
   
         Total2(13) = Total2(13) + Amt - CostAmount
         Total3(13) = Total3(13) + Amt - CostAmount
         Total4(13) = Total4(13) + Amt - CostAmount
      Next m_BillingDoc
         
      Call GenerateFooter(PrevKey2, Total4, m_Details1)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      CountColor = CountColor + 1
      If CountColor > 2 Then
         Vsp.TextColor = RGB(0, 150, 0)
         If CountColor > 3 Then
            CountColor = 0
         End If
      End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      Vsp.TextColor = RGB(0, 0, 0)
   
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
      
   'รายละเอียดจุลินทรีย์
   Vsp.TextColor = RGB(0, 0, 0)
   Vsp = ""
   Vsp = "รายละเอียดจุลินทรีย์"
   Call printHeader2
      
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
      Total4(j) = 0
      Total7(j) = 0
   Next j
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
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   m_BillingDoc.APAR_IND = 1
   Call m_BillingDoc.QueryData(157, Rs, iCount)
      
   Call GetFirstLastDate(mcolParam("FROM_BILL_DATE"), FromDate, ToDate)
   TempDateCountMonth = DateDiff("D", FromDate, ToDate) + 1
   TempFirstDayOfMonth = FromDate
      
   While Not Rs.EOF
      Set m_BillingDoc = New CBillingDoc
      Call m_BillingDoc.PopulateFromRS2(157, Rs)
         
      If m_BillingDoc.STOCK_NO = "58-1017" Then
         I = 2
         TempFirstDayOfMonth = FromDate
         For j = 1 To TempDateCountMonth
            If Weekday(TempFirstDayOfMonth) = 1 Or j = TempDateCountMonth Then
               I = I + 2   'นับสัปดาห์
            End If
            If Day(m_BillingDoc.DOCUMENT_DATE) = Day(TempFirstDayOfMonth) Then
               If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
                  Total1(I) = Total1(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total1(I + 1) = Total1(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
                  Total1(I) = Total1(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total1(I + 1) = Total1(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               End If
            End If
            TempFirstDayOfMonth = DateAdd("d", 1, TempFirstDayOfMonth)
         Next j
        ElseIf m_BillingDoc.STOCK_NO = "58-1018" Then
         I = 2
         TempFirstDayOfMonth = FromDate
         For j = 1 To TempDateCountMonth
            If Weekday(TempFirstDayOfMonth) = 1 Or j = TempDateCountMonth Then
               I = I + 2   'นับสัปดาห์
            End If
            If Day(m_BillingDoc.DOCUMENT_DATE) = Day(TempFirstDayOfMonth) Then
               If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
                  Total4(I) = Total4(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total4(I + 1) = Total4(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
                  Total4(I) = Total4(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total4(I + 1) = Total4(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               End If
            End If
            TempFirstDayOfMonth = DateAdd("d", 1, TempFirstDayOfMonth)
         Next j
      ElseIf m_BillingDoc.STOCK_NO = "58-1009" Then
         I = 2
         TempFirstDayOfMonth = FromDate
         For j = 1 To TempDateCountMonth
            If Weekday(TempFirstDayOfMonth) = 1 Or j = TempDateCountMonth Then
                  I = I + 2   'นับสัปดาห์
            End If
            If Day(m_BillingDoc.DOCUMENT_DATE) = Day(TempFirstDayOfMonth) Then
               If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
                  Total2(I) = Total2(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total2(I + 1) = Total2(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
                  Total2(I) = Total2(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total2(I + 1) = Total2(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               End If
            End If
               TempFirstDayOfMonth = DateAdd("d", 1, TempFirstDayOfMonth)
            Next j
      ElseIf m_BillingDoc.STOCK_NO = "58-1020" Then
         I = 2
         TempFirstDayOfMonth = FromDate
         For j = 1 To TempDateCountMonth
            If Weekday(TempFirstDayOfMonth) = 1 Or j = TempDateCountMonth Then
                  I = I + 2   'นับสัปดาห์
            End If
            If Day(m_BillingDoc.DOCUMENT_DATE) = Day(TempFirstDayOfMonth) Then
               If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
                  Total7(I) = Total2(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total7(I + 1) = Total2(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) + m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) + m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
                  Total7(I) = Total2(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total7(I + 1) = Total2(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
                  Total3(I) = Total3(I) - m_BillingDoc.TOTAL_AMOUNT
                  Total3(I + 1) = Total3(I + 1) - m_BillingDoc.TOTAL_PRICE - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT
               End If
            End If
               TempFirstDayOfMonth = DateAdd("d", 1, TempFirstDayOfMonth)
            Next j
         End If

      Rs.MoveNext
   Wend
   Call m_Details1.ClearField

   j = TempFinalPosition - 2
   While (j > 0)
      Total1(TempFinalPosition) = Total1(TempFinalPosition) + Total1(j)
      Total1(TempFinalPosition + 1) = Total1(TempFinalPosition + 1) + Total1(j - 1)
      Total2(TempFinalPosition) = Total2(TempFinalPosition) + Total2(j)
      Total2(TempFinalPosition + 1) = Total2(TempFinalPosition + 1) + Total2(j - 1)
      Total3(TempFinalPosition) = Total3(TempFinalPosition) + Total3(j)
      Total3(TempFinalPosition + 1) = Total3(TempFinalPosition + 1) + Total3(j - 1)
      Total4(TempFinalPosition) = Total4(TempFinalPosition) + Total4(j)
      Total4(TempFinalPosition + 1) = Total4(TempFinalPosition + 1) + Total4(j - 1)
      Total7(TempFinalPosition) = Total7(TempFinalPosition) + Total7(j)
      Total7(TempFinalPosition + 1) = Total7(TempFinalPosition + 1) + Total7(j - 1)
      j = j - 2
   Wend
   
   Call GenerateFooterMicroorganism("ซานิไซม์-200พรีมิกซ์ 40กก/58-1018", Total4, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
      
   Call GenerateFooterMicroorganism("ซานิไซม์-200พรีมิกซ์ 20กก/58-1020", Total7, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   
   Call GenerateFooterMicroorganism("ซานิไซม์(กก.)/58-1017", Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
         
   Call GenerateFooterMicroorganism("มอสน็อค(6กรัม/ซอง)/58-1009", Total2, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   
   Call GenerateFooterMicroorganism("รวม ", Total3, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TextColor = RGB(0, 0, 255)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   Vsp.TextColor = RGB(0, 0, 0)
   'จบรายงานที่ 1
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       "รายงานยอดขาย แยกตามพนักงานขาย รวมตามกลุ่มวันที่" & vbCrLf & _
                                       "จากงวด " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A") & " " & "ถึง " & EmptyToString(DateToStringExtEx2(DateAdd("D", DateCount - 1, mcolParam("FROM_DATE"))), "N/A") & " รายวัน"
   CheckPage = 2
   mblnHeader = False
   Vsp.NewPage
   Call printHeader3
   mblnHeader = True
   
   For j = 1 To UBound(Total2)
      Total2(j) = 0
      Total4(j) = 0
   Next j
   I = 0
   CountColor = 0
   
   For Each m_BillingDoc In m_DistinctMain
'      If m_BillingDoc.SALE_CODE <> "" Then
         I = I + 1
         If PrevKey1 <> m_BillingDoc.SALE_CODE And I <> 1 Then
            k = 1
            For Each Tg In m_TagetDetails
               If Tg.Flag <> "Y" And Tg.EMPLOYEE_CODE = PrevKey1 Then
                  k = k + 1
                  Total2(k) = Total2(k) + Tg.TOTAL_PRICE
                  Total4(k) = Total4(k) + Tg.TOTAL_PRICE
                  'เป้ารายวัน
                  If TempDateCountNow <> TempDateCountTagetDetails Then
                     k = k + 1
                     Total2(k) = Total2(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
                     Total4(k) = Total4(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
                  End If
                  k = 1
               End If
            Next Tg
            If Total2(TempFirstPosition) > 0 Then
                  Total2(TempLastPosition - 1) = -1 * ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total2(TempLastPosition - 2))
                  Total2(TempLastPosition) = (Total2(TempLastPosition - 2) * 100) / ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
            End If
            Call GenerateFooter2(PrevKey2 & " " & PrevKey1, Total2, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            CountColor = CountColor + 1
            If CountColor > 2 Then
               Vsp.TextColor = RGB(0, 0, 0)
               If CountColor > 3 Then
                  CountColor = 0
               End If
            Else
               Vsp.TextColor = RGB(0, 150, 0)
            End If
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.TextColor = RGB(0, 0, 0)
            
            For j = 1 To UBound(Total2)
               Total2(j) = 0
            Next j
         ElseIf I = 1 Then
            
         End If

         PrevKey1 = m_BillingDoc.SALE_CODE
         PrevKey2 = m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME

         Call m_Details1.ClearField
         k = 1
         'เป้าการขายเดือน
         Amt = 0
         Set Tg = GetObject("CTagetDetail", m_TagetDetails, m_BillingDoc.SALE_CODE & "-" & m_BillingDoc.STOCK_NO)
         k = k + 1
         TempFirstPosition = k
         Amt = Tg.TOTAL_PRICE
         If Not Tg.Flag = "Y" Then
         Else
            Amt = 0
         End If
         Total2(k) = Total2(k) + Amt
         Total4(k) = Total4(k) + Amt
         'เป้าการขาย วัน
         If TempDateCountNow <> TempDateCountTagetDetails Then
            k = k + 1
            Amt = (Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails
            If Not Tg.Flag = "Y" Then
            Else
               Amt = 0
            End If
            Total2(k) = Total2(k) + Amt
            Total4(k) = Total4(k) + Amt
         End If
         If Amt <> 0 Then
            Tg.Flag = "Y"
         End If
         
         TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
         TempDateString = Trim(Replace(TempDateString, ")", ""))
         TempDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
         TotalPrice = 0

         While Len(TempDateString) > 0
            Call CalculateDatePeriod(TempDateString, CountDate)
            k = k + 1
            Amt = 0
            While CountDate > 0
               TempDate = DateAdd("D", 1, TempDate)
             
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
               Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
               Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
               Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)

               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
               Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
               Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
               Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
               'F&B
               If m_BillingDoc.JOINT_CODE <> "" Then
                  Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
                  Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
                  Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
                  Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)

                  Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
                  Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
                  Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
                  Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
               End If
               CountDate = CountDate - 1
            Wend
            TotalPrice = TotalPrice + Amt
            Total2(k) = Total2(k) + Amt
            Total4(k) = Total4(k) + Amt
         Wend
         k = k + 1
         Total2(k) = Total2(k) + TotalPrice
         Total4(k) = Total4(k) + TotalPrice
         'ผลต่าง
         k = k + 1
         '%
         k = k + 1
         TempLastPosition = k
'      End If
   Next m_BillingDoc
      
   k = 1
   For Each Tg In m_TagetDetails
      If Tg.Flag <> "Y" And Tg.EMPLOYEE_CODE = PrevKey1 Then
         k = k + 1
         Total2(k) = Total2(k) + Tg.TOTAL_PRICE
         Total4(k) = Total4(k) + Tg.TOTAL_PRICE
         'เป้ารายวัน
         If TempDateCountNow <> TempDateCountTagetDetails Then
            k = k + 1
            Total2(k) = Total2(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
            Total4(k) = Total4(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
         End If
         k = 1
      End If
   Next Tg
      
   If Total2(TempFirstPosition) > 0 Then
      Total2(TempLastPosition - 1) = -1 * ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total2(TempLastPosition - 2))
      Total2(TempLastPosition) = (Total2(TempLastPosition - 2) * 100) / ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
   End If
   Call GenerateFooter2(PrevKey2 & " " & PrevKey1, Total2, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   CountColor = CountColor + 1
   If CountColor > 2 Then
      Vsp.TextColor = RGB(0, 0, 0)
      If CountColor > 3 Then
         CountColor = 0
      End If
   Else
      Vsp.TextColor = RGB(0, 150, 0)
   End If
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.TextColor = RGB(0, 0, 0)
      
   If Total4(TempFirstPosition) > 0 Then
      Total4(TempLastPosition - 1) = (Total4(TempLastPosition - 2) - (Total4(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
      Total4(TempLastPosition) = (Total4(TempLastPosition - 2) * 100) / ((Total4(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
   End If
   Call GenerateFooter2("      สุทธิ" & "" & "", Total4, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TextColor = RGB(0, 0, 255)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   Vsp.TextColor = RGB(0, 0, 0)
   'ยอดขายรายวัน
   Vsp = ""
   Vsp = ""
   Vsp = "ยอดขาย"
   Call printHeader6
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   For Each m_BillingDoc In m_DistinctMain
      k = 1
      TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
      TempDateString = Trim(Replace(TempDateString, ")", ""))
      TempDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
      TotalPrice = 0
      TotalPrice2 = 0
      
      While Len(TempDateString) > 0
         Call CalculateDatePeriod(TempDateString, CountDate)
         k = k + 1
         Amt = 0
         Amt2 = 0
         While CountDate > 0
            TempDate = DateAdd("D", 1, TempDate)
             
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
            Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)

            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
            Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
            'F&B
            If m_BillingDoc.JOINT_CODE <> "" Then
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)

               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
            End If
            CountDate = CountDate - 1
         Wend
         TotalPrice = TotalPrice + Amt
         TotalPrice2 = TotalPrice2 + Amt2
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt2
      Wend
      k = k + 1
      Total1(k) = Total1(k) + TotalPrice
      Total2(k) = Total2(k) + TotalPrice2
   Next m_BillingDoc
   
   Call GenerateFooter4("QMC", Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid - 4000, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)

   Call GenerateFooter4("F&B", Total2, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid - 4000, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   'จบรายงานที่2
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       "รายงานเปรียบเทียบเป้าการขาย ยอดขาย แยกตามพนักงานขาย รายวัน สะสม" & vbCrLf & _
                                        "จากพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
   
   CheckPage = 3
   mblnHeader = False
   Vsp.NewPage
   Call printHeader4
   mblnHeader = True
   
   Call GetFirstLastDate(mcolParam("TO_BILL_DATE"), FromDate, ToDate)
   TempDateCountRepost3 = DateDiff("D", FromDate, mcolParam("TO_BILL_DATE")) + 1 - Val(mcolParam("HOLIDAY2"))

   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
   I = 0
   CountColor = 0
   Call GetFirstLastDate(mcolParam("TO_BILL_DATE"), FromDate, ToDate)
   For Each m_BillingDoc In m_DistinctMain2
      I = I + 1
      Amt = 0
      Call m_Details1.ClearField

      Set HeadCf = m_Headers4.Fields(1)
      TempStr = m_BillingDoc.SALE_CODE & " /  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      'เป้ารายวัน
      Set HeadCf = m_Headers4.Fields(2)
      Set Tg = GetObject("CTagetDetail", m_TagetDetails2, m_BillingDoc.SALE_CODE)
      TempStr = FormatNumberToNull(Tg.TOTAL_PRICE / TempDateCountTaget, 0)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(2) = Total1(2) + (Tg.TOTAL_PRICE / TempDateCountTaget)
      'ยอดรายวัน
      Set HeadCf = m_Headers4.Fields(3)
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
         
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
      'F&B
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
         
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
      
      If Amt < 0 Then
         TempStr = "(" & FormatNumberToNull(-1 * Amt, 0) & ")"
      Else
         TempStr = FormatNumberToNull(Amt, 0)
      End If
      Total1(3) = Total1(3) + Amt
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      'Diff รายวัน
      Set HeadCf = m_Headers4.Fields(4)
      If Amt - (Tg.TOTAL_PRICE / TempDateCountTaget) < 0 Then
         TempStr = "(" & FormatNumberToNull(-1 * (Amt - (Tg.TOTAL_PRICE / TempDateCountTaget)), 0) & ")"
      Else
         TempStr = FormatNumberToNull(Amt - ((Tg.TOTAL_PRICE / TempDateCountTaget)), 0)
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(4) = Total1(4) + (Amt - (Tg.TOTAL_PRICE / TempDateCountTaget))
      '% รายวัน
      Set HeadCf = m_Headers4.Fields(5)
      If Tg.TOTAL_PRICE > 0 Then
         TempStr = FormatNumberToNull((Amt * 100) / (Tg.TOTAL_PRICE / TempDateCountTaget))
      Else
         TempStr = ""
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      'เป้าสะสม
      Set HeadCf = m_Headers4.Fields(6)
      TempStr = FormatNumberToNull((Tg.TOTAL_PRICE * TempDateCountRepost3) / TempDateCountTaget, 0)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(6) = Total1(6) + ((Tg.TOTAL_PRICE * TempDateCountRepost3) / TempDateCountTaget)
      'ยอดสะสม
      Amt = 0
      Set HeadCf = m_Headers4.Fields(7)
      TempDate = FromDate
      While TempDate <= ToDate
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
            
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
         'F&B
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
            
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
         
         TempDate = DateAdd("D", 1, TempDate)
      Wend
      If Amt < 0 Then
         TempStr = "(" & FormatNumberToNull(-1 * Amt, 0) & ")"
      Else
         TempStr = FormatNumberToNull(Amt, 0)
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(7) = Total1(7) + Amt
      'Diff สะสม
      Set HeadCf = m_Headers4.Fields(8)
      If Amt - Tg.TOTAL_PRICE < 0 Then
         TempStr = "(" & FormatNumberToNull(-1 * (Amt - (Tg.TOTAL_PRICE * TempDateCountRepost3) / TempDateCountTaget), 0) & ")"
      Else
         TempStr = FormatNumberToNull((Amt - (Tg.TOTAL_PRICE * TempDateCountRepost3) / TempDateCountTaget), 0)
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      Total1(8) = Total1(8) + (Amt - (Tg.TOTAL_PRICE * TempDateCountRepost3) / TempDateCountTaget)
      '% สะสม
      Set HeadCf = m_Headers4.Fields(9)
      If Tg.TOTAL_PRICE > 0 Then
         TempStr = FormatNumberToNull((Amt * 100) / ((Tg.TOTAL_PRICE * TempDateCountRepost3) / TempDateCountTaget))
      Else
         TempStr = ""
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
         
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      CountColor = CountColor + 1
      If CountColor > 2 Then
         Vsp.TextColor = RGB(0, 0, 0)
         If CountColor > 3 Then
            CountColor = 0
         End If
      Else
         Vsp.TextColor = RGB(0, 150, 0)
      End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.TextColor = RGB(0, 0, 0)
   Next m_BillingDoc
      
   If Total1(2) <> 0 Then
      Total1(5) = ((Total1(3) * 100) / Total1(2))
   End If
   If Total1(6) <> 0 Then
      Total1(9) = ((Total1(7) * 100) / Total1(6))
   End If
   
   Call GenerateFooter3("      สุทธิ", Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Vsp.TextColor = RGB(0, 0, 255)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   Vsp.TextColor = RGB(0, 0, 0)
   
   'ยอดขายรวม
'   Vsp = ""
'   Vsp = ""
'   Vsp = "ยอดขายรวม"
'   Call printHeader5
'   For j = 1 To UBound(Total1)
'      Total1(j) = 0
'      Total2(j) = 0
'   Next j
'   For Each m_BillingDoc In m_FirstDistinctMain2
'      Amt = 0
'      Call m_Details1.ClearField
'      'ยอดรายวัน
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'      Total1(2) = Total1(2) + Amt
'      'ยอดสะสม
'      Amt = 0
'      TempDate = FromDate
'      While TempDate <= ToDate
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
'         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs2, m_BillingDoc.SALE_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
'         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'         TempDate = DateAdd("D", 1, TempDate)
'      Wend
'      Total1(3) = Total1(3) + Amt
'   Next m_BillingDoc
'   'F&B
'   For Each m_BillingDoc In m_DistinctMain_FB2
'      Amt = 0
'      Call m_Details1.ClearField
'      'ยอดรายวัน
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'      Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & mcolParam("TO_BILL_DATE"))
'      Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'      Total2(2) = Total2(2) + Amt
'      'ยอดสะสม
'      Amt = 0
'      TempDate = FromDate
'      While TempDate <= ToDate
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmounts_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
'         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
'         Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
'         Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs_FB2, m_BillingDoc.SALE_JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
'         Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
'
'         TempDate = DateAdd("D", 1, TempDate)
'      Wend
'      Total2(3) = Total2(3) + Amt
'   Next m_BillingDoc
'
'   Call GenerateFooter4("QMC", Total1, m_Details1)
'   Call m_Details1.GetString(1, TempStr1, TempStr2)
'   strFormat = VSP_CalTable(TempStr1, mdblWid - 8000, alngX)
'   Call Vsp.AddTable(strFormat, "", TempStr2)
'
'   Call GenerateFooter4("F&B", Total2, m_Details1)
'   Call m_Details1.GetString(1, TempStr1, TempStr2)
'   strFormat = VSP_CalTable(TempStr1, mdblWid - 8000, alngX)
'   Call Vsp.AddTable(strFormat, "", TempStr2)
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''
   Vsp = ""
   Vsp = ""
   Vsp = "ยอดขาย"
   Call printHeader6
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   For Each m_BillingDoc In m_DistinctMain
      k = 1
      TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
      TempDateString = Trim(Replace(TempDateString, ")", ""))
      TempDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
      TotalPrice = 0
      TotalPrice2 = 0
      
      While Len(TempDateString) > 0
         Call CalculateDatePeriod(TempDateString, CountDate)
         k = k + 1
         Amt = 0
         Amt2 = 0
         While CountDate > 0
            TempDate = DateAdd("D", 1, TempDate)
             
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
            Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)

            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
            Amt = Amt + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
            Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
            Amt = Amt - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
            'F&B
            If m_BillingDoc.JOINT_CODE <> "" Then
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmounts_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)

               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & INVOICE_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 + TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
               Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountExs_FB, m_BillingDoc.SALE_JOINT_CODE & "-" & m_BillingDoc.JOINT_CODE & "-" & RETURN_DOCTYPE & "-" & TempDate)
               Amt2 = Amt2 - (TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
            End If
            CountDate = CountDate - 1
         Wend
         TotalPrice = TotalPrice + Amt
         TotalPrice2 = TotalPrice2 + Amt2
         Total1(k) = Total1(k) + Amt
         Total2(k) = Total2(k) + Amt2
      Wend
      k = k + 1
      Total1(k) = Total1(k) + TotalPrice
      Total2(k) = Total2(k) + TotalPrice2
   Next m_BillingDoc
   
   Call GenerateFooter4("QMC", Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid - 4000, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)

   Call GenerateFooter4("F&B", Total2, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid - 4000, alngX)
   Call Vsp.AddTable(strFormat, "", TempStr2)
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''
   
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set m_BillingDoc = Nothing
   Set TempBillingDoc = Nothing
   Set TempBillingDoc2 = Nothing
   Set TempBillingDoc3 = Nothing
   Set TempBillingDocFB = Nothing
   Set TempBillingDocFB2 = Nothing
   Set TempBillingDocFB3 = Nothing

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
         TempStr = FormatNumberToNull(Tot(j), 0)
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j), 0) & ")"
         End If
      Else
         TempStr = FormatNumberToNull(MyDiff(Tot(j + 1), Tot(j - 1)))
         If MyDiff(Tot(j + 1), Tot(j - 1)) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * MyDiff(Tot(j + 1), Tot(j - 1))) & ")"
         End If
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
         TempStr = FormatNumberToNull(Tot(j), 0)
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j), 0) & ")"
         End If
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter3(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   j = 0
   For Each HeadCf In m_Headers4.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf (j >= 2 And j < 5) Or (j >= 6 And j < 9) Then
         TempStr = FormatNumberToNull(Tot(j), 0)
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j), 0) & ")"
         End If
      ElseIf j = 5 Or j = 9 Then
         TempStr = FormatNumberToNull(Tot(j))
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j)) & ")"
         End If
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
   
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter4(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   j = 0
   For Each HeadCf In m_Headers5.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j >= 2 Then
         TempStr = FormatNumberToNull(Tot(j))
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j)) & ")"
         End If
      End If
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
   
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooterMicroorganism(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempAmount As Double
Dim TempSale As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers3.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      Else
         TempStr = FormatNumberToNull(Tot(j), 0)
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j), 0) & ")"
         End If
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
Private Sub GenerateFooter2(txt1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField

   j = 0
   For Each HeadCf In m_Headers4.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j >= 2 And j < (TempLastPosition - 1) Then
         TempStr = FormatNumberToNull(Tot(j), 0)
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j), 0) & ")"
         End If
      ElseIf j = TempLastPosition - 1 Then
         TempStr = FormatNumberToNull(Tot(j))
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j), 0) & ")"
         End If
       ElseIf j = TempLastPosition Then
         TempStr = FormatNumberToNull(Tot(j))
         If Tot(j) < 0 Then
            TempStr = "(" & FormatNumberToNull(-1 * Tot(j)) & ")"
         End If
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
      If CheckPage = 2 Then
         Call printHeader3
      ElseIf CheckPage = 3 Then
         Call printHeader4
      Else
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
'   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
'   End If
End Sub
Private Function GetKey1(Tg As CBillingDoc) As String
   GetKey1 = Trim(Tg.SALE_CODE & "-" & Tg.STOCK_NO)
End Function
Private Sub SumDistinctStockNos(Cl As Collection, Cl2 As Collection, Cl3 As Collection)
Dim m_BillingDoc As CBillingDoc
Dim m_BillingDoc2 As CBillingDoc
Dim TempData As CBillingDoc
Dim TempCheck As String
Dim TempCode As String
Dim I As Long
Dim FoundFlag As Boolean
   
   Set Cl3 = Nothing
   Set Cl3 = New Collection
   For Each m_BillingDoc In Cl
      If Cl2.Count <> 0 Then
         If m_BillingDoc.STOCK_NO = "56-6005" Or m_BillingDoc.STOCK_NO = "11-1006" Then
            Debug.Print
         End If
         For Each m_BillingDoc2 In Cl2
            If m_BillingDoc2.STOCK_NO = "56-6005" Or m_BillingDoc2.STOCK_NO = "11-1006" Then
               Debug.Print
            End If
            If m_BillingDoc.JOINT_CODE = m_BillingDoc2.STOCK_NO Then
               m_BillingDoc2.Flag = "Y"
            End If
         Next m_BillingDoc2
         Call Cl3.add(m_BillingDoc, Trim(m_BillingDoc.STOCK_GROUP_CODE & "-" & m_BillingDoc.STOCK_NO))
      Else
         Call Cl3.add(m_BillingDoc)
      End If
   Next m_BillingDoc
   
   'ตรวจยอดคงเหลือ
   For Each m_BillingDoc2 In Cl2
      If m_BillingDoc2.STOCK_NO = "56-6005" Or m_BillingDoc2.STOCK_NO = "11-1006" Then
         Debug.Print
      End If
      If m_BillingDoc2.Flag <> "Y" Then
         TempCode = m_BillingDoc2.STOCK_NO
         m_BillingDoc2.STOCK_NO = m_BillingDoc2.JOINT_CODE
         m_BillingDoc2.JOINT_CODE = TempCode
         m_BillingDoc2.Flag = "Y"
         Set TempData = GetObject("CBillingDoc", Cl3, Trim(m_BillingDoc2.STOCK_GROUP_CODE & "-" & m_BillingDoc2.JOINT_CODE), False)
         If TempData Is Nothing Then
            I = 0
            FoundFlag = False
            For Each TempData In Cl3
               I = I + 1
               If TempData.STOCK_GROUP_CODE = m_BillingDoc2.STOCK_GROUP_CODE Then
                  FoundFlag = True
                  Exit For
               End If
            Next TempData
            If FoundFlag Then
               Call Cl3.add(m_BillingDoc2, Trim(m_BillingDoc2.STOCK_GROUP_CODE & "-" & m_BillingDoc2.JOINT_CODE), I)
            Else
               Call Cl3.add(m_BillingDoc2, Trim(m_BillingDoc2.STOCK_GROUP_CODE & "-" & m_BillingDoc2.JOINT_CODE), , I)
            End If
         End If
      End If
   Next m_BillingDoc2
End Sub
Private Sub CalculateDate(TempStr As String, DateCount As Long)
Dim TempID As Long
Dim TempStrNew  As String
   TempStrNew = Replace(TempStr, "(", "")
   TempStrNew = Replace(TempStrNew, ")", "")
   TempID = InStr(1, TempStrNew, ",")
   While InStr(1, TempStrNew, ",") > 0
      TempID = InStr(1, TempStrNew, ",")
      DateCount = DateCount + Val(Left(TempStrNew, TempID - 1))   ' 7+7+7+7+3 =31
      TempStrNew = Mid(TempStrNew, TempID + 1)       ' ตัดวันที่ ที่ใช้ไปแล้ว
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
Private Sub SumDistinctMain(Cl As Collection, Cl2 As Collection, Cl3 As Collection)
Dim m_BillingDoc As CBillingDoc
Dim m_BillingDoc2 As CBillingDoc
Dim TempData As CBillingDoc
Dim TempKey1 As String
Dim TempCode As String

   Set Cl3 = Nothing
   Set Cl3 = New Collection
   TempKey1 = ""
   For Each m_BillingDoc In Cl
      'เปลี่ยน sale ตรวจสอบว่ามี sale F&B รหัสก่อนหน้า sale QMC หรือไม่ ถ้ามีก็เพิ่มลงไป
      If TempKey1 <> m_BillingDoc.SALE_CODE Then
         For Each m_BillingDoc2 In Cl2
            If m_BillingDoc.SALE_CODE > m_BillingDoc2.SALE_JOINT_CODE And m_BillingDoc2.Flag <> "Y" Then
               TempCode = m_BillingDoc2.JOINT_CODE
               m_BillingDoc2.JOINT_CODE = m_BillingDoc2.STOCK_NO
               m_BillingDoc2.STOCK_NO = TempCode
               TempCode = m_BillingDoc2.SALE_JOINT_CODE
               m_BillingDoc2.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE
               m_BillingDoc2.SALE_CODE = TempCode
               m_BillingDoc2.Flag = "Y"
               Call Cl3.add(m_BillingDoc2)
            End If
         Next m_BillingDoc2
      End If
      'เจอรายการน้ำของ sale QMC ตรวจสอบว่ามีตรงกับ ของ F&Bหรือไม่ หากเจอ ให้ set flag="Y"
      If m_BillingDoc.JOINT_CODE <> "" Then
         For Each m_BillingDoc2 In Cl2
            If m_BillingDoc.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE And m_BillingDoc.JOINT_CODE = m_BillingDoc2.STOCK_NO And m_BillingDoc2.Flag <> "Y" Then
               m_BillingDoc2.Flag = "Y"
            End If
         Next m_BillingDoc2
         Call Cl3.add(m_BillingDoc)
         'เติมน้ำอีกรายการของsaleคนเดิม ต่อจากน้ำตัวบน
         For Each m_BillingDoc2 In Cl2
            TempCode = ""
            If m_BillingDoc.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE And m_BillingDoc.JOINT_CODE <> m_BillingDoc2.STOCK_NO And m_BillingDoc2.Flag <> "Y" Then
               TempCode = m_BillingDoc2.JOINT_CODE
               m_BillingDoc2.JOINT_CODE = m_BillingDoc2.STOCK_NO
               m_BillingDoc2.STOCK_NO = TempCode
               TempCode = m_BillingDoc2.SALE_JOINT_CODE
               m_BillingDoc2.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE
               m_BillingDoc2.SALE_CODE = TempCode
               m_BillingDoc2.Flag = "Y"
               Call Cl3.add(m_BillingDoc2)
            End If
         Next m_BillingDoc2
      Else
         Call Cl3.add(m_BillingDoc)
      End If
      TempKey1 = m_BillingDoc.SALE_JOINT_CODE
   Next m_BillingDoc
   'เติมรายการที่เหลือ
   For Each m_BillingDoc2 In Cl2
      TempCode = ""
      If m_BillingDoc2.Flag <> "Y" Then
         TempCode = m_BillingDoc2.JOINT_CODE
         m_BillingDoc2.JOINT_CODE = m_BillingDoc2.STOCK_NO
         m_BillingDoc2.STOCK_NO = TempCode
         TempCode = m_BillingDoc2.SALE_JOINT_CODE
         m_BillingDoc2.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE
         m_BillingDoc2.SALE_CODE = TempCode
         m_BillingDoc2.Flag = "Y"
         Call Cl3.add(m_BillingDoc2)
      End If
   Next m_BillingDoc2
   
End Sub
Private Sub SumDistinctMain2(Cl As Collection, Cl2 As Collection, Cl3 As Collection)
Dim m_BillingDoc As CBillingDoc
Dim m_BillingDoc2 As CBillingDoc
Dim TempData As CBillingDoc
Dim TempKey1 As String
Dim TempCode As String

   Set Cl3 = Nothing
   Set Cl3 = New Collection
   TempKey1 = ""
   For Each m_BillingDoc In Cl
     If TempKey1 <> m_BillingDoc.SALE_CODE Then
         For Each m_BillingDoc2 In Cl2
            If m_BillingDoc.SALE_CODE > m_BillingDoc2.SALE_JOINT_CODE And m_BillingDoc2.Flag <> "Y" Then
               TempCode = m_BillingDoc2.SALE_JOINT_CODE
               m_BillingDoc2.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE
               m_BillingDoc2.SALE_CODE = TempCode
               m_BillingDoc2.Flag = "Y"
               Call Cl3.add(m_BillingDoc2)
            End If
         Next m_BillingDoc2
      End If
      If m_BillingDoc.SALE_JOINT_CODE <> "" Then
         For Each m_BillingDoc2 In Cl2
            If m_BillingDoc.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE And m_BillingDoc2.Flag <> "Y" Then
               m_BillingDoc2.Flag = "Y"
            End If
         Next m_BillingDoc2
         Call Cl3.add(m_BillingDoc)
      Else
         Call Cl3.add(m_BillingDoc)
      End If
      TempKey1 = m_BillingDoc.SALE_CODE
   Next m_BillingDoc
   'เติมที่เหลือ
   For Each m_BillingDoc2 In Cl2
      TempCode = ""
      If m_BillingDoc2.Flag <> "Y" Then
         TempCode = m_BillingDoc2.SALE_JOINT_CODE
         m_BillingDoc2.SALE_JOINT_CODE = m_BillingDoc2.SALE_CODE
         m_BillingDoc2.SALE_CODE = TempCode
         m_BillingDoc2.Flag = "Y"
         Call Cl3.add(m_BillingDoc2)
      End If
   Next m_BillingDoc2
End Sub
Private Sub SumTempDistinctMain(Cl As Collection, Cl2 As Collection, Cl3 As Collection)
Dim m_BillingDoc As CBillingDoc
Dim m_BillingDoc2 As CBillingDoc
Dim m_TagetDetail As CTagetDetail
Dim TempData As CBillingDoc
Dim TempTagetDetail As CTagetDetail
Dim TempKey1 As String
Dim TempKey2 As String

   Set Cl3 = Nothing
   Set Cl3 = New Collection
   TempKey1 = ""
   TempKey2 = ""
   For Each m_BillingDoc In Cl
      If TempKey1 <> m_BillingDoc.SALE_CODE Then
         For Each m_TagetDetail In Cl2
            If m_BillingDoc.SALE_CODE > m_TagetDetail.EMPLOYEE_CODE And m_TagetDetail.Flag <> "Y" Then
               Set m_BillingDoc2 = New CBillingDoc
               m_BillingDoc2.STOCK_NO = m_TagetDetail.STOCK_NO
               m_BillingDoc2.JOINT_CODE = m_TagetDetail.JOINT_CODE
               m_BillingDoc2.STOCK_DESC = m_TagetDetail.STOCK_DESC
               m_BillingDoc2.UNIT_AMOUNT = m_TagetDetail.UNIT_AMOUNT
               m_BillingDoc2.SALE_CODE = m_TagetDetail.EMPLOYEE_CODE
               m_BillingDoc2.SALE_JOINT_CODE = m_TagetDetail.SALE_JOINT_CODE
               m_BillingDoc2.SALE_LONG_NAME = m_TagetDetail.SALE_LONG_NAME
               m_BillingDoc2.SALE_LAST_NAME = m_TagetDetail.SALE_LAST_NAME
               m_TagetDetail.Flag = "Y"
               Call Cl3.add(m_BillingDoc2)
            End If
         Next m_TagetDetail
      Else
         For Each m_TagetDetail In Cl2
            If m_BillingDoc.SALE_CODE >= m_TagetDetail.EMPLOYEE_CODE And m_BillingDoc.STOCK_NO > m_TagetDetail.STOCK_NO And m_TagetDetail.Flag <> "Y" Then
               Set m_BillingDoc2 = New CBillingDoc
               m_BillingDoc2.STOCK_NO = m_TagetDetail.STOCK_NO
               m_BillingDoc2.JOINT_CODE = m_TagetDetail.JOINT_CODE
               m_BillingDoc2.STOCK_DESC = m_TagetDetail.STOCK_DESC
               m_BillingDoc2.UNIT_AMOUNT = m_TagetDetail.UNIT_AMOUNT
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
         If m_BillingDoc.SALE_CODE = m_TagetDetail.EMPLOYEE_CODE And m_BillingDoc.STOCK_NO = m_TagetDetail.STOCK_NO And m_TagetDetail.Flag <> "Y" Then
            m_TagetDetail.Flag = "Y"
         End If
      Next m_TagetDetail
      Call Cl3.add(m_BillingDoc)
      TempKey1 = m_BillingDoc.SALE_CODE
      TempKey2 = m_BillingDoc.STOCK_NO
   Next m_BillingDoc
   'เติมที่เหลือ
   For Each m_TagetDetail In Cl2
      If m_TagetDetail.Flag <> "Y" Then
         Set m_BillingDoc2 = New CBillingDoc
         m_BillingDoc2.STOCK_NO = m_TagetDetail.STOCK_NO
         m_BillingDoc2.JOINT_CODE = m_TagetDetail.JOINT_CODE
         m_BillingDoc2.STOCK_DESC = m_TagetDetail.STOCK_DESC
         m_BillingDoc2.UNIT_AMOUNT = m_TagetDetail.UNIT_AMOUNT
         m_BillingDoc2.SALE_CODE = m_TagetDetail.EMPLOYEE_CODE
         m_BillingDoc2.SALE_JOINT_CODE = m_TagetDetail.SALE_JOINT_CODE
         m_BillingDoc2.SALE_LONG_NAME = m_TagetDetail.SALE_LONG_NAME
         m_BillingDoc2.SALE_LAST_NAME = m_TagetDetail.SALE_LAST_NAME
         Call Cl3.add(m_BillingDoc2)
      End If
   Next m_TagetDetail
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

