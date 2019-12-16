VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportCommission001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportCommission001"
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
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

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
Private FromDate As Date
Private ToDate As Date
Private m_TotalSum As Collection
Private m_TotalReturn As Collection
Private m_ComTablePrice As Collection
Private m_ComTableAmount As Collection
Private m_ComTableReturn As Collection
Private m_TotalTop As Collection
Private CommissionChart As Collection
Private SumTotalChart As Collection
Private m_Taget As Collection
Private m_ComTableEx As Collection
Private m_ComTableBd As Collection
Private m_GroupComColl As Collection
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
   
   Set m_TotalSum = New Collection
   Set m_TotalReturn = New Collection

   Set m_ComTablePrice = New Collection
   Set m_ComTableAmount = New Collection
   Set m_ComTableReturn = New Collection
   Set m_TotalTop = New Collection
   Set CommissionChart = New Collection
   Set SumTotalChart = New Collection
   Set m_Taget = New Collection
   Set m_ComTableEx = New Collection
   Set m_ComTableBd = New Collection
   Set m_GroupComColl = New Collection
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
   
   Set m_TotalSum = Nothing
   Set m_TotalReturn = Nothing
   
   Set m_ComTablePrice = Nothing
   Set m_ComTableAmount = Nothing
   Set m_ComTableReturn = Nothing
   Set m_TotalTop = Nothing
   Set CommissionChart = Nothing
   Set SumTotalChart = Nothing
   Set m_Taget = Nothing
   Set m_ComTableEx = Nothing
   Set m_ComTableBd = Nothing
   Set m_GroupComColl = Nothing
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
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(1, "^", "NO", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "กลุ่มคอมมิตชั่น", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เป้าการขาย จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เป้าการขาย บาท", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดขาย จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดขาย บาท", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดขาย (จำนวน)TOP", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดขาย (บาท)TOP", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รวมยอดขาย (จำนวน)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รวมยอดขาย (บาท)", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "%ยอดขาย จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "%ยอดขาย บาท", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ค่าCom.", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "% ยอดคืน", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดขายหัก CN (บาท)", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ค่าCom ที่ได้รับ", ">")
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
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 1440
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call GetFirstLastDate(DateSerial(mcolParam("YEAR_NO") - 543, mcolParam("MONTH_ID"), 1), FromDate, ToDate)
      
   Call LoadCommissionTotalSum(m_TotalSum, FromDate, ToDate, "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")")
   Call LoadCommissionTotalSum(m_TotalReturn, FromDate, ToDate, "(" & RETURN_DOCTYPE & ")")
   
   Call LoadCommissionTabelPrice(m_ComTablePrice, FromDate, ToDate, COMMISSION_TABLE)
   Call LoadCommissionTabelPrice(m_ComTableReturn, FromDate, ToDate, RETURN_TABLE)
   
   Call LoadSumTop(Nothing, m_TotalTop, FromDate, ToDate)
   
   Call LoadTagetByEmpGroupcom(m_Taget, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")))
   
   Call LoadCommissionTabelEx(m_ComTableEx, FromDate, ToDate, COMMISSION_TABLE_EX)
      
   Call LoadMaster(Nothing, m_GroupComColl, , , MASTER_GROUP_COM)
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        IntToThaiMonth(mcolParam("MONTH_ID")) & " " & mcolParam("YEAR_NO")
   
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
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As Long
Dim PrevKey2 As Long
Dim PrevKey3 As Long
Dim PrevKey4 As Boolean
Dim CB As CBillingDoc
Dim Rb As CBillingDoc
Dim Plb As CPrintLabel
Dim TotalCom As CTotalCommission
Dim TotalComCol As Collection
Dim m_CommissionChart As CCommissionChart
Dim TotalCommission As CTotalCommission
Dim TempComAmount As Double
Dim TempReturn As Double
Dim Taget As CTagetDetail
Dim CanSum As Boolean
Dim TempMr As CMasterRef
Dim HaveValue As Boolean

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
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      Set BodyCf = New CReportField

      Set m_CommissionChart = New CCommissionChart
      Set TotalCom = New CTotalCommission
      
      Call m_CommissionChart.SetFieldValue("COMMISSION_CHART_ID", -1)
      Call m_CommissionChart.SetFieldValue("FROM_COMMISSION_DATE", FromDate)
      Call m_CommissionChart.SetFieldValue("TO_COMMISSION_DATE", FromDate)
      Call m_CommissionChart.SetFieldValue("ORDER_TYPE", 1)
      
      Call m_CommissionChart.QueryData(4, Rs, iCount)

      If iCount <= 0 Then
         mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
         Exit Function
      End If
      
      While Not Rs.EOF
         Set m_CommissionChart = New CCommissionChart
         Call m_CommissionChart.PopulateFromRS(4, Rs)
         Call CommissionChart.add(m_CommissionChart, Trim(Str(m_CommissionChart.GetFieldValue("COMMISSION_CHART_ID"))))
         Rs.MoveNext
      Wend
      Call GenerateComEx
      Call SumChart(CommissionChart, SumTotalChart)
'      Call printHeader
      I = 0
      
      If Rs.EOF Then
         Rs.MoveFirst
      End If

      While Not Rs.EOF
         I = I + 1
         Call m_CommissionChart.PopulateFromRS(4, Rs)
         CanSum = CheckSum(m_CommissionChart.GetFieldValue("EMP_ID"))

         If I > 1 Then
            Call GenerateFooter("", "รวม", "", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            Vsp = m_CommissionChart.GetFieldValue("EMP_NAME") & " " & m_CommissionChart.GetFieldValue("EMP_LNAME")
         ElseIf I = 1 Then
            Vsp = m_CommissionChart.GetFieldValue("EMP_NAME") & " " & m_CommissionChart.GetFieldValue("EMP_LNAME")
         End If

         For Each TempMr In m_GroupComColl
            HaveValue = False
            Call m_Details1.ClearField
            
            TempComAmount = 0
              
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = TempMr.KEY_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set TotalCommission = GetObject("CTotalCommission", SumTotalChart, Trim(m_CommissionChart.GetFieldValue("EMP_ID") & "-" & TempMr.KEY_ID))
            
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = FormatNumberToNull(TotalCommission.TAGET_AMOUNT_SUM)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(3) = Total2(3) + TotalCommission.TAGET_AMOUNT_SUM
            End If
            Total1(3) = Total1(3) + TotalCommission.TAGET_AMOUNT_SUM
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TAGET_AMOUNT_SUM)
            
            Set HeadCf = m_Headers1.Fields(4)
            TempStr = FormatNumberToNull(TotalCommission.TAGET_PRICE_SUM)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(4) = Total2(4) + TotalCommission.TAGET_PRICE_SUM
            End If
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TAGET_PRICE_SUM)
            Total1(4) = Total1(4) + TotalCommission.TAGET_PRICE_SUM
               
            Set HeadCf = m_Headers1.Fields(5)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_AMOUNT_SUM)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(5) = Total2(5) + TotalCommission.TOTAL_AMOUNT_SUM
            End If
            Total1(5) = Total1(5) + TotalCommission.TOTAL_AMOUNT_SUM
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TOTAL_AMOUNT_SUM)
            
            Set HeadCf = m_Headers1.Fields(6)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_PRICE_SUM)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(6) = Total2(6) + TotalCommission.TOTAL_PRICE_SUM
            End If
            Total1(6) = Total1(6) + TotalCommission.TOTAL_PRICE_SUM
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TOTAL_PRICE_SUM)
            
            Set HeadCf = m_Headers1.Fields(7)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_AMOUNT_TOP)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(7) = Total2(7) + TotalCommission.TOTAL_AMOUNT_TOP
            End If
            Total1(7) = Total1(7) + TotalCommission.TOTAL_AMOUNT_TOP
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TOTAL_AMOUNT_TOP)
            
            Set HeadCf = m_Headers1.Fields(8)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_PRICE_TOP)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(8) = Total2(8) + TotalCommission.TOTAL_PRICE_TOP
            End If
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TOTAL_PRICE_TOP)
            Total1(8) = Total1(8) + TotalCommission.TOTAL_PRICE_TOP
            
            Set HeadCf = m_Headers1.Fields(9)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(9) = Total2(9) + TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP
            End If
            Total1(9) = Total1(9) + TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP
            
            Set HeadCf = m_Headers1.Fields(10)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(10) = Total2(10) + TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP
            End If
            Total1(10) = Total1(10) + TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP
            
            Set HeadCf = m_Headers1.Fields(11)
            If TotalCommission.TAGET_AMOUNT_SUM > 0 Then
               TempStr = MyDiff((TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP) * 100, TotalCommission.TAGET_AMOUNT_SUM)
            Else
               TempStr = "-"
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(12)
            If TotalCommission.TAGET_PRICE_SUM > 0 Then
               TempStr = MyDiff((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * 100, TotalCommission.TAGET_PRICE_SUM)
            Else
               TempStr = "-"
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(13)
            TempStr = FormatNumberToNull(Trim(CalculateCom(m_CommissionChart.GetFieldValue("EMP_ID"), TotalCommission, TempMr.KEY_ID)))
            TempComAmount = Val(Replace(TempStr, ",", ""))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(13) = Total1(13) + TempComAmount
            Total2(13) = Total2(13) + TempComAmount
            
            Set HeadCf = m_Headers1.Fields(14)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_PRICE_RETURN)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            If CanSum Then
               Total2(14) = Total2(14) + TotalCommission.TOTAL_PRICE_RETURN
            End If
            Total1(14) = Total1(14) + TotalCommission.TOTAL_PRICE_RETURN
            HaveValue = CheckHaveValue(HaveValue, TotalCommission.TOTAL_PRICE_RETURN)
            
            Set HeadCf = m_Headers1.Fields(15)
            If TotalCommission.TOTAL_PRICE_SUM > 0 Then
               TempStr = MyDiff(TotalCommission.TOTAL_PRICE_RETURN * 100, TotalCommission.TOTAL_PRICE_SUM)
            Else
               TempStr = "-"
            End If
            TempReturn = Val(Replace(TempStr, ",", ""))
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(16)
            TempStr = FormatNumberToNull(TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP - TotalCommission.TOTAL_PRICE_RETURN)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(16) = Total1(16) + TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP - TotalCommission.TOTAL_PRICE_RETURN
            Total2(16) = Total2(16) + TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP - TotalCommission.TOTAL_PRICE_RETURN
            
            Set HeadCf = m_Headers1.Fields(17)
            If TempComAmount > 0 Then
               TempStr = FormatNumber(Trim(Str(CalculateReturn(m_CommissionChart.GetFieldValue("EMP_ID"), TotalCommission, TempComAmount, TempMr.KEY_ID))), 0)
            Else
               TempStr = 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(17) = Total1(17) + Val(Replace(TempStr, ",", ""))
            Total2(17) = Total2(17) + Val(Replace(TempStr, ",", ""))
            
            If HaveValue Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         Next TempMr
         Rs.MoveNext
      Wend

      Call GenerateFooter("", "รวม", "", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("", "สุทธิ", "", Total2, m_Details1, 0)
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

Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
      ElseIf j = 11 Then
         TempStr = FormatNumberToNull(MyDiff(Tot(9) * 100, Tot(3)))
      ElseIf j = 12 Then
         TempStr = FormatNumberToNull(MyDiff(Tot(10) * 100, Tot(4)))
      ElseIf j = 15 Then
         TempStr = FormatNumberToNull(MyDiff(Tot(14) * 100, Tot(6)))
      ElseIf j = m_Headers1.Fields.Count Then
         TempStr = FormatNumberToNull(Tot(j), 0)
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
Private Sub SumChart(Coll As Collection, SumTotal As Collection)
Dim Cm As CCommissionChart
Dim ParentID As Long
Dim TempMr As CMasterRef
   For Each Cm In Coll
      For Each TempMr In m_GroupComColl
         Call Recuresive(Cm, SumTotal, GetParent(Cm.GetFieldValue("COMMISSION_CHART_ID")), GetEmp(Cm.GetFieldValue("COMMISSION_CHART_ID")), TempMr.KEY_ID)
      Next TempMr
   Next Cm
End Sub
Public Sub Recuresive(Cm As CCommissionChart, SumTotal As Collection, ParentID As Long, OwnId As Long, ComGroupId As Long)
On Error Resume Next
Dim P1 As CTotalCommission
Dim P2 As CTotalCommission
Set P1 = New CTotalCommission
Dim Old As Double
Dim CB As CBillingDoc
Dim Plb As CPrintLabel
Dim Rb As CBillingDoc
Dim Tg As CTagetDetail
Dim Mfe As CMasterFromToEx
   P1.EMP_ID = OwnId
   P1.GROUP_COM_ID = ComGroupId
   
   Set CB = GetObject("CBillingDoc", m_TotalSum, Trim(Cm.GetFieldValue("EMP_ID") & "-" & ComGroupId))
   Set Plb = GetObject("CPrintLabel", m_TotalTop, Trim(Cm.GetFieldValue("EMP_ID") & "-" & ComGroupId))
   Set Rb = GetObject("CBillingDoc", m_TotalReturn, Trim(Cm.GetFieldValue("EMP_ID") & "-" & ComGroupId))
   Set Tg = GetObject("CTaget", m_Taget, Trim(Cm.GetFieldValue("EMP_ID") & "-" & ComGroupId))
   Set Mfe = GetObject("CMasterFromToEx", m_ComTableBd, Trim(Cm.GetFieldValue("EMP_ID") & "-" & ComGroupId))
   
   P1.TOTAL_AMOUNT_SUM = CB.TOTAL_AMOUNT
   'debug.print (Cm.GetFieldValue("EMP_NAME") & "-" & Mfe.PERCENT)
   P1.TOTAL_PRICE_SUM = CB.TOTAL_PRICE - CB.DISCOUNT_AMOUNT - CB.EXT_DISCOUNT_AMOUNT + Mfe.PERCENT
   
   P1.TOTAL_PRICE_RETURN = Rb.TOTAL_PRICE - Rb.DISCOUNT_AMOUNT - Rb.EXT_DISCOUNT_AMOUNT
   
   P1.TOTAL_AMOUNT_TOP = Plb.GetFieldValue("TOTAL_AMOUNT")
   P1.TOTAL_PRICE_TOP = Plb.GetFieldValue("TOTAL_PRICE") - Plb.GetFieldValue("DISCOUNT_AMOUNT") - Plb.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
   
   P1.TAGET_AMOUNT_SUM = Tg.GetFieldValue("TOTAL_AMOUNT")
   P1.TAGET_PRICE_SUM = Tg.GetFieldValue("TOTAL_PRICE")
   If SumTotal.Count = 0 Then
      Call SumTotal.add(P1, Trim(P1.Getkey))
   Else
      Set P2 = SumTotal(Trim(P1.Getkey))
      If P2 Is Nothing Then
         Call SumTotal.add(P1, Trim(P1.Getkey))
      Else
         P2.TOTAL_AMOUNT_SUM = P2.TOTAL_AMOUNT_SUM + P1.TOTAL_AMOUNT_SUM
         P2.TOTAL_PRICE_SUM = P2.TOTAL_PRICE_SUM + P1.TOTAL_PRICE_SUM
         
         P2.TOTAL_PRICE_RETURN = P2.TOTAL_PRICE_RETURN + P1.TOTAL_PRICE_RETURN
         
         P2.TOTAL_AMOUNT_TOP = P2.TOTAL_AMOUNT_TOP + P1.TOTAL_AMOUNT_TOP
         P2.TOTAL_PRICE_TOP = P2.TOTAL_PRICE_TOP + P1.TOTAL_PRICE_TOP
         
         P2.TAGET_AMOUNT_SUM = P2.TAGET_AMOUNT_SUM + P1.TAGET_AMOUNT_SUM
         P2.TAGET_PRICE_SUM = P2.TAGET_PRICE_SUM + P1.TAGET_PRICE_SUM
      End If
   End If
   
   If ParentID > 0 Then
      Call Recuresive(Cm, SumTotal, GetParent(ParentID), GetEmp(ParentID), ComGroupId)
   End If
End Sub
Private Function GetParent(ID As Long) As Long
Dim Cm As CCommissionChart
   Set Cm = GetObject("CCommissionChart", CommissionChart, Trim(Str(ID)))
   GetParent = Cm.GetFieldValue("PARENT_ID")
End Function
Private Function GetEmp(ID As Long) As Long
Dim Cm As CCommissionChart
   Set Cm = GetObject("CCommissionChart", CommissionChart, Trim(Str(ID)))
   GetEmp = Cm.GetFieldValue("EMP_ID")
End Function
Private Function CalculateCom(EMP_ID As Long, TotalCommission As CTotalCommission, GroupComID As Long) As Double
Dim Cc As CCommissionChart
Dim ChartID As Long
Dim Mfd As CMasterFromToDetail
Dim TempCom As Double
   For Each Cc In CommissionChart
      If Cc.GetFieldValue("EMP_ID") = EMP_ID Then
         ChartID = Cc.GetFieldValue("SALE_ID")
         Exit For
      End If
   Next Cc
   Set Cc = Nothing
   
   TempCom = 0
   For Each Mfd In m_ComTablePrice
      ''debug.print (Mfd.GetFieldValue("MASTER_FROMTO_ID"))
      If Mfd.GetFieldValue("MASTER_FROMTO_ID") = ChartID And Mfd.GetFieldValue("GROUP_COM_ID") = GroupComID Then
         '-----------------------------------------------------------------------
         If Mfd.GetFieldValue("MULTIPLE_FLAG") = "Y" Then
            TempCom = MyDiffEx((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * Mfd.GetFieldValue("MULTIPLE_PERCENT"), 100)
            If Mfd.GetFieldValue("INCLUDE_SUB_FLAG") = "Y" Then
               TempCom = MyDiff(TempCom * Mfd.GetFieldValue("INCLUDE_SUB_PERCENT"), 100)
            End If
            Exit Function
         End If
         '-----------------------------------------------------------------------
         If Mfd.GetFieldValue("STEP_FLAG") = "Y" Then
            If Mfd.GetFieldValue("AMOUNT_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= Round(TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP, 2) And Round(TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP, 2) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubStepComm(Mfd, TotalCommission)
               End If
            ElseIf Mfd.GetFieldValue("VALUE_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= Round((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP), 2) And Round((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP), 2) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubStepComm(Mfd, TotalCommission)
               End If
            ElseIf Mfd.GetFieldValue("TAGET_VALUE_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= MyDiff((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * 100, TotalCommission.TAGET_PRICE_SUM) And MyDiff((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * 100, TotalCommission.TAGET_PRICE_SUM) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubStepComm(Mfd, TotalCommission)
               End If
            End If
         End If
      '-----------------------------------------------------------------------
         If Mfd.GetFieldValue("TIER_FLAG") = "Y" Then
            If Mfd.GetFieldValue("AMOUNT_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= Round(TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP, 2) And Round(TotalCommission.TOTAL_AMOUNT_SUM + TotalCommission.TOTAL_AMOUNT_TOP, 2) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubTierComm(Mfd, TotalCommission)
               End If
            ElseIf Mfd.GetFieldValue("VALUE_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= Round((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP), 2) And Round((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP), 2) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubTierComm(Mfd, TotalCommission)
               End If
            ElseIf Mfd.GetFieldValue("TAGET_VALUE_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= MyDiff((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * 100, TotalCommission.TAGET_PRICE_SUM) And MyDiff((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * 100, TotalCommission.TAGET_PRICE_SUM) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubTierComm(Mfd, TotalCommission)
               End If
            End If
         End If
         '-----------------------------------------------------------------------
      End If
   Next Mfd
      
   CalculateCom = TempCom
End Function
Private Function CalculateSubStepComm(Mfd As CMasterFromToDetail, TotalCommission As CTotalCommission) As Double
Dim TempCom As Double
   If Mfd.GetFieldValue("NO_X_FLAG") = "Y" Then
      TempCom = Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE2")
   ElseIf Mfd.GetFieldValue("AMOUNT_X_FLAG") = "Y" Then
      TempCom = MyDiffEx(TotalCommission.TOTAL_AMOUNT_SUM * Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"), 100)
   ElseIf Mfd.GetFieldValue("VALUE_X_FLAG") = "Y" Then
      TempCom = MyDiffEx((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) * Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"), 100)
   End If
   If Mfd.GetFieldValue("INCLUDE_SUB_FLAG") = "Y" Then
      TempCom = MyDiff(TempCom * Mfd.GetFieldValue("INCLUDE_SUB_PERCENT"), 100)
   End If
   CalculateSubStepComm = TempCom
End Function
Private Function CalculateSubTierComm(Mfd As CMasterFromToDetail, TotalCommission As CTotalCommission) As Double
Dim TempCom As Double
   If Mfd.GetFieldValue("NO_X_FLAG") = "Y" Then
      TempCom = Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE2")
   ElseIf Mfd.GetFieldValue("AMOUNT_X_FLAG") = "Y" Then
      TempCom = MyDiffEx((TotalCommission.TOTAL_AMOUNT_SUM - Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM")) * Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"), 100)
   ElseIf Mfd.GetFieldValue("VALUE_X_FLAG") = "Y" Then
      TempCom = MyDiffEx(((TotalCommission.TOTAL_PRICE_SUM + TotalCommission.TOTAL_PRICE_TOP) - Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM")) * Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"), 100)
   End If
   TempCom = TempCom + Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE3")
   If Mfd.GetFieldValue("INCLUDE_SUB_FLAG") = "Y" Then
      TempCom = MyDiff(TempCom * Mfd.GetFieldValue("INCLUDE_SUB_PERCENT"), 100)
   End If
   CalculateSubTierComm = TempCom
End Function
Private Function CalculateReturn(EMP_ID As Long, TotalCommission As CTotalCommission, Commission As Double, GroupComID As Long) As Double
Dim Cc As CCommissionChart
Dim ChartID As Long
Dim Mfd As CMasterFromToDetail
Dim TempCom As Double
   For Each Cc In CommissionChart
      If Cc.GetFieldValue("EMP_ID") = EMP_ID Then
         ChartID = Cc.GetFieldValue("RETURN_ID")
         Exit For
      End If
   Next Cc
   Set Cc = Nothing
   
   TempCom = Commission
   For Each Mfd In m_ComTableReturn
      If Mfd.GetFieldValue("MASTER_FROMTO_ID") = ChartID And Mfd.GetFieldValue("GROUP_COM_ID") = GroupComID Then
         '-----------------------------------------------------------------------
         If Mfd.GetFieldValue("MULTIPLE_FLAG") = "Y" Then
            TempCom = MyDiff(Commission * Mfd.GetFieldValue("MULTIPLE_PERCENT"), 100)
            If Mfd.GetFieldValue("INCLUDE_SUB_FLAG") = "Y" Then
               TempCom = MyDiff(TempCom * Mfd.GetFieldValue("INCLUDE_SUB_PERCENT"), 100)
            End If
            Exit For
         End If
         '-----------------------------------------------------------------------
         If Mfd.GetFieldValue("STEP_FLAG") = "Y" Then
            If Mfd.GetFieldValue("VALUE_P_FLAG") = "Y" Then
               If Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_FROM") <= MyDiff(TotalCommission.TOTAL_PRICE_RETURN * 100, TotalCommission.TOTAL_PRICE_SUM) And MyDiff(TotalCommission.TOTAL_PRICE_RETURN * 100, TotalCommission.TOTAL_PRICE_SUM) <= Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_TO") Then
                  TempCom = CalculateSubStepReturn(Mfd, TotalCommission, Commission)
                  Exit For
               End If
            End If
         End If
      '-----------------------------------------------------------------------
      End If
   Next Mfd
      
   CalculateReturn = TempCom
End Function
Private Function CalculateSubStepReturn(Mfd As CMasterFromToDetail, TotalCommission As CTotalCommission, Commission As Double) As Double
Dim TempCom As Double
   If Mfd.GetFieldValue("NO_X_FLAG") = "Y" Then
      TempCom = Commission - Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE2")
   ElseIf Mfd.GetFieldValue("VALUE_X_FLAG") = "Y" Then
      TempCom = MyDiffEx(Commission * Mfd.GetFieldValue("MASTER_FROMTO_DETAIL_VALUE1"), 100)
   End If
   If Mfd.GetFieldValue("INCLUDE_SUB_FLAG") = "Y" Then
      TempCom = MyDiff(TempCom * Mfd.GetFieldValue("INCLUDE_SUB_PERCENT"), 100)
   End If
   CalculateSubStepReturn = TempCom
End Function
Private Sub GenerateComEx()
Dim Mfe As CMasterFromToEx
Dim PrevKey1  As String
Dim PrevKey2  As String
Dim PrevKey3  As String
Dim CustomerSet As String
Dim PartItemSet As String
Dim EmpSet As String
Dim EmpAmountSet As String
Dim EmpNameSet As String
   
   For Each Mfe In m_ComTableEx
      If Trim(Str(Mfe.MASTER_FROMTO_ID)) <> PrevKey1 And Len(PrevKey1) > 0 Then
         '------------------------------------------------------------------------------------------
         Call GenerateComExAmount(CustomerSet, PartItemSet, EmpSet, EmpAmountSet, EmpNameSet, PrevKey1, PrevKey2, PrevKey3)
         
      End If
      PrevKey1 = Trim(Str(Mfe.MASTER_FROMTO_ID))
      PrevKey2 = Trim(Mfe.MASTER_FROMTO_NO)
      PrevKey3 = Trim(Mfe.MASTER_FROMTO_DESC)
         
      If Mfe.CUSTOMER_ID > 0 And Len(CustomerSet) <= 0 Then
         CustomerSet = "(" & Mfe.CUSTOMER_ID & ","
      ElseIf Mfe.CUSTOMER_ID > 0 Then
         CustomerSet = CustomerSet & Mfe.CUSTOMER_ID & ","
      End If
      
      If Mfe.PART_ITEM_ID > 0 And Len(PartItemSet) <= 0 Then
         PartItemSet = "(" & Mfe.PART_ITEM_ID & ","
      ElseIf Mfe.PART_ITEM_ID > 0 Then
         PartItemSet = PartItemSet & Mfe.PART_ITEM_ID & ","
      End If
      
      If Mfe.EMP_ID > 0 And Len(EmpSet) <= 0 Then
         EmpSet = "(" & Mfe.EMP_ID & ","
         EmpAmountSet = "(" & Mfe.PERCENT & ","
         EmpNameSet = "(" & Mfe.SALE_NAME & "(" & Mfe.SALE_CODE & ")" & ","
         
      ElseIf Mfe.EMP_ID > 0 Then
         EmpSet = EmpSet & Mfe.EMP_ID & ","
         EmpAmountSet = EmpAmountSet & Mfe.PERCENT & ","
         EmpNameSet = EmpNameSet & Mfe.SALE_NAME & "(" & Mfe.SALE_CODE & ")" & ","
      End If
         
   Next Mfe
   
   Call GenerateComExAmount(CustomerSet, PartItemSet, EmpSet, EmpAmountSet, EmpNameSet, PrevKey1, PrevKey2, PrevKey3)
   '------------------------------------------------------------------------------------------
   
   
End Sub
Public Sub GenerateComExAmount(CustomerSet As String, PartItemSet As String, EmpSet As String, EmpAmountSet As String, EmpNameSet As String, PrevKey1 As String, PrevKey2 As String, PrevKey3 As String)
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim TempRs As ADODB.Recordset
Dim TempBd As CMasterFromToEx
Dim TempMr As CMasterRef
   
   CustomerSet = Left(CustomerSet, Len(CustomerSet) - 1) & ")"
   PartItemSet = Left(PartItemSet, Len(PartItemSet) - 1) & ")"
   EmpSet = Left(EmpSet, Len(EmpSet) - 1) & ")"
   EmpAmountSet = Left(EmpAmountSet, Len(EmpAmountSet) - 1) & ")"
   EmpNameSet = Left(EmpNameSet, Len(EmpNameSet) - 1) & ")"
   '------------------------------------------------------------------------------------------
   
   Set TempRs = New ADODB.Recordset
   
   Set BD = New CBillingDoc
   BD.APAR_CODE = CustomerSet
   BD.STOCK_NO = PartItemSet
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")"
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryData(96, TempRs, ItemCount)
   
   While Not TempRs.EOF
      Call BD.PopulateFromRS2(96, TempRs)
      
      Set TempMr = GetObject("CMasterRef", m_GroupComColl, Trim(Str(BD.GROUP_COM_ID)))
      Vsp = "กลุ่ม " & TempMr.KEY_NAME
      Vsp = "ประเภทคอมมิตชั่นส่วนเพิ่ม   " & PrevKey3 & " ( " & PrevKey2 & " )" & " ยอดขาย " & FormatNumber(BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT)
      
      If Len(EmpSet) > 0 Then
         EmpSet = Mid(EmpSet, 2, Len(EmpSet) - 2)
         EmpAmountSet = Mid(EmpAmountSet, 2, Len(EmpAmountSet) - 2)
         EmpNameSet = Mid(EmpNameSet, 2, Len(EmpNameSet) - 2)
      End If
      While Len(EmpSet) > 0
         If InStr(1, EmpSet, ",") > 0 Then
                     
            Vsp = "         พนักงานขาย  " & Left(EmpNameSet, InStr(1, EmpNameSet, ",") - 1) & "   ยอดขาย   " & FormatNumber((BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT) * Val(Left(EmpAmountSet, InStr(1, EmpAmountSet, ",") - 1)) / 100)
            
            Set TempBd = GetObject("CMasterFromToEx", m_ComTableBd, Trim(Left(EmpSet, InStr(1, EmpSet, ",") - 1) & "-" & BD.GROUP_COM_ID), False)
            If TempBd Is Nothing Then
               Set TempBd = New CMasterFromToEx
               TempBd.EMP_ID = Left(EmpSet, InStr(1, EmpSet, ",") - 1)
               TempBd.SALE_NAME = Left(EmpNameSet, InStr(1, EmpNameSet, ",") - 1)
               TempBd.PERCENT = (BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT) * Val(Left(EmpAmountSet, InStr(1, EmpAmountSet, ",") - 1)) / 100
               TempBd.GROUP_COM_ID = BD.GROUP_COM_ID
               Call m_ComTableBd.add(TempBd, Trim(TempBd.EMP_ID & "-" & BD.GROUP_COM_ID))
            Else
               TempBd.PERCENT = TempBd.PERCENT + (BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT) * Val(Left(EmpAmountSet, InStr(1, EmpAmountSet, ",") - 1)) / 100
            End If
            
            Set TempBd = Nothing
            
            EmpSet = Right(EmpSet, Len(EmpSet) - InStr(1, EmpSet, ","))
            EmpAmountSet = Right(EmpAmountSet, Len(EmpAmountSet) - InStr(1, EmpAmountSet, ","))
            EmpNameSet = Right(EmpNameSet, Len(EmpNameSet) - InStr(1, EmpNameSet, ","))
         Else
            
            Vsp = "         พนักงานขาย  " & EmpNameSet & "   ยอดขาย   " & FormatNumber((BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT) * Val(EmpAmountSet) / 100)
            
            Set TempBd = GetObject("CMasterFromToEx", m_ComTableBd, EmpSet, False)
            If TempBd Is Nothing Then
               Set TempBd = New CMasterFromToEx
               TempBd.EMP_ID = EmpSet
               TempBd.SALE_NAME = EmpNameSet
               TempBd.PERCENT = (BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT) * Val(EmpAmountSet) / 100
               TempBd.GROUP_COM_ID = BD.GROUP_COM_ID
               Call m_ComTableBd.add(TempBd, Trim(TempBd.EMP_ID & "-" & TempBd.GROUP_COM_ID))
            Else
               TempBd.PERCENT = TempBd.PERCENT + (BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT) * Val(EmpAmountSet) / 100
            End If
            Set TempBd = Nothing
            
            EmpSet = ""
            EmpAmountSet = ""
            EmpNameSet = ""
            
            Vsp = ""
         End If
      Wend
      TempRs.MoveNext
   Wend
      
   CustomerSet = ""
   PartItemSet = ""
   EmpSet = ""
   EmpAmountSet = ""
   EmpNameSet = ""
         
   Set TempRs = Nothing
End Sub
Private Function CheckSum(ID As Long) As Boolean
Dim Mfe As CCommissionChart
   CheckSum = True
   For Each Mfe In CommissionChart
      If Mfe.GetFieldValue("PARENT_ID") > 0 And ID = Mfe.GetFieldValue("EMP_ID") Then
         CheckSum = False
         Exit Function
      End If
   Next Mfe
End Function

