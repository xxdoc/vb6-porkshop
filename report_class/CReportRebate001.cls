VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportRebate001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportRebate001"
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
Private TempDateCountTaget As Long
Private TempFromDate  As Date
Private TempToDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<52|<4|<8;"

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

Private m_SaleAmounts As Collection
Private m_SaleAmountsByCus As Collection

Private SumTotalChart As Collection
Private NetworkChartColl As Collection
Private OrderNetworkChartColl As Collection
Private CollRebateCondition As Collection
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
   
   Set m_SaleAmounts = New Collection
   Set m_SaleAmountsByCus = New Collection
   Set CollRebateCondition = New Collection
   
   Set SumTotalChart = New Collection
   Set NetworkChartColl = New Collection
   Set OrderNetworkChartColl = New Collection
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
   
   Set m_SaleAmounts = Nothing
   Set m_SaleAmountsByCus = Nothing
   
   Set SumTotalChart = Nothing
   Set NetworkChartColl = Nothing
   Set OrderNetworkChartColl = Nothing
   
   Set CollRebateCondition = Nothing
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
Dim TempFromDate  As Date
Dim TempToDate   As Date

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   Call Cf.SetFieldValue(2, "^", "รหัส", "<")
   Call m_Headers1.AddField(Cf)

   Call Cf.SetFieldValue(4, "^", "ลูกค้า", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "REBATE", ">", "SUM_FLAG")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "สรุป REBATE", ">", "SUM_FLAG")
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
Dim D As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 500
   Vsp.MarginFooter = 500
   Vsp.MarginHeader = 500
   Vsp.MarginLeft = 500
   Vsp.MarginRight = 500
   Vsp.MarginTop = 500
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   Vsp.ZoomMode = zmPageWidth
   
   Call GetSaleAmountByCustomer(m_SaleAmountsByCus, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   Call GetSaleAmountCustomerStockGroup(m_SaleAmounts, mcolParam("FROM_DATE"), mcolParam("TO_DATE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("FROM_CUSTOMER_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   Call LoadDataRebateCondition(CollRebateCondition, mcolParam("FROM_DATE"), mcolParam("TO_DATE"))
   
   FromDate = "จากงวด " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " " & "จากลูกค้า " & EmptyToString(mcolParam("FROM_CUSTOMER_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_CUSTOMER_CODE"), "N/A") & _
                                       "จากสินค้า" & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A")
                                       
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
Dim TempBd As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim TempDate As Date
Dim TempBillingDoc As CBillingDoc

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

      Set BodyCf = New CReportField
      I = 0
      
      Dim m_NetworkChart As CNetworkChart
      Dim TotalSale As CTotalSale
      
      Set m_NetworkChart = New CNetworkChart
      Set TotalSale = New CTotalSale
      
      m_NetworkChart.MASTER_VALID_ID = -1
      m_NetworkChart.VALID_FROM = mcolParam("FROM_DATE")
      m_NetworkChart.VALID_TO = mcolParam("TO_DATE")
      Call m_NetworkChart.QueryData(2, Rs, iCount, True)

      If iCount <= 0 Then
         mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
         Exit Function
      End If

      While Not Rs.EOF
         Set m_NetworkChart = New CNetworkChart
         Call m_NetworkChart.PopulateFromRS(2, Rs)
         Call NetworkChartColl.add(m_NetworkChart, Trim(Str(m_NetworkChart.NETWORK_CHART_ID)))
         Rs.MoveNext
      Wend
      
      Call GenerateOrderNetworkChart(NetworkChartColl, -1, 0)
      
      Call SumChart(OrderNetworkChartColl, SumTotalChart)
      
      For Each m_NetworkChart In OrderNetworkChartColl
      
         Call m_Details1.ClearField
   
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_NetworkChart.CUSTOMER_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = m_NetworkChart.CUSTOMER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set TotalSale = GetObject("CTotalSale", SumTotalChart, Trim(Str(m_NetworkChart.CUSTOMER_ID)))
                  
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = FormatNumberToNull(TotalSale.TOTAL_PRICE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(3) = Total1(3) + (TotalSale.TOTAL_PRICE)
         
         Set HeadCf = m_Headers1.Fields(4)
         Set TempBillingDoc = GetObject("CBillingDoc", m_SaleAmountsByCus, Trim(Str(m_NetworkChart.CUSTOMER_ID)))
         If TempBillingDoc.TOTAL_PRICE > 0 Then       'มียอดขายก็ให้ค่า Rebate ไม่มียอดขายก็ไม่ให้ค่า Rebate
            If m_NetworkChart.HOLD_FLAG = "Y" Then 'ไม่ได้ค่า Rebate เนื่องจากไม่มีส่วนร่วม หรือ บริจาคเข้าองค์กรกุศล
               TempStr = "*"
               Total1(4) = Total1(4) + (0)
            Else
               TempStr = FormatNumberToNull(TotalSale.TOTAL_PRICE)
               Total1(4) = Total1(4) + (TotalSale.TOTAL_PRICE)
            End If
         Else
            If m_NetworkChart.HOLD_FLAG = "Y" Then 'ไม่ได้ค่า Rebate เนื่องจากไม่มีส่วนร่วม หรือ บริจาคเข้าองค์กรกุศล
               TempStr = "*" & "," & "***"
               Total1(4) = Total1(4) + (0)
            Else
               TempStr = "***"
               Total1(4) = Total1(4) + (0)
            End If
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         
         
         If mcolParam("SHOW_COLOR") = 1 Then
            If m_NetworkChart.Level = 0 Then
               Vsp.TextColor = RGB(0, 0, 0)
            ElseIf m_NetworkChart.Level = 1 Then
               Vsp.TextColor = RGB(255, 0, 0)
            ElseIf m_NetworkChart.Level = 2 Then
               Vsp.TextColor = RGB(0, 0, 255)
            ElseIf m_NetworkChart.Level = 3 Then
               Vsp.TextColor = RGB(0, 255, 0)
            Else
               Vsp.TextColor = RGB(((m_NetworkChart.Level + 1) * 200) Mod 256, (m_NetworkChart.Level * 50) Mod 256, (m_NetworkChart.Level * 90) Mod 256)
            End If
         End If
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Next m_NetworkChart
      
      Vsp.TextColor = RGB(0, 0, 0)
      
      Call GenerateFooter("      สุทธิ" & "" & "", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Set m_SaleAmounts = Nothing
      
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

Private Sub GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
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
      ElseIf HeadCf.TempKey = "SUM_FLAG" Then
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
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage & " / " & glbUser.USER_NAME)
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
   Vsp = "*** ไม่ได้ค่า REBATE เนื่องจากไม่มียอดขาย"
   Vsp = "* ไม่ได้ค่า REBATE เนื่องจากระงับการจ่าย หรือ ไม่มีส่วนร่วมในแผนธุรกิจ หรือ ไม่ขอรับ  Rebate"
End Sub
Private Sub SumChart(Coll As Collection, SumTotal As Collection)
Dim Cm As CNetworkChart
Dim TempBd As CBillingDoc
Dim BD As CBillingDoc
   For Each Cm In Coll
      Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Trim(Str(Cm.CUSTOMER_ID)), False)
      If Not TempBd Is Nothing Then
         For Each BD In TempBd.CollBillingDoc
            Call Recuresive(Cm, SumTotal, GetParent(Cm.NETWORK_CHART_ID), GetCustomer(Cm.NETWORK_CHART_ID), 0, BD)
         Next BD
      End If
   Next Cm
End Sub
Public Sub Recuresive(Cm As CNetworkChart, SumTotal As Collection, ParentID As Long, OwnId As Long, Level As Long, BD As CBillingDoc)
On Error Resume Next
Dim Amt As Double
Dim P1 As CTotalSale
Dim P2 As CTotalSale
Set P1 = New CTotalSale
Dim Old As Double
Dim TempBd As CBillingDoc
Dim RbCdt As CRebateCondition
   
   P1.CUSTOMER_ID = OwnId
   P1.CUSTOMER_NAME = Cm.CUSTOMER_NAME & " (" & Cm.CUSTOMER_CODE & ")"
   
   Amt = 0
   If BD.DOCUMENT_TYPE = INVOICE_DOCTYPE Or BD.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
      'Amt = Amt + BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT
      Amt = Amt + BD.TOTAL_PRICE
   ElseIf BD.DOCUMENT_TYPE = RETURN_DOCTYPE Then
      'Amt = Amt - (BD.TOTAL_PRICE - BD.DISCOUNT_AMOUNT - BD.EXT_DISCOUNT_AMOUNT)
      Amt = Amt - (BD.TOTAL_PRICE)
   End If
   
   'set ค่าของประเภทลูกค้า และ กลุ่มลูกค้า และ LEVEL เพื่อหา Percent ที่จะคูณกับ Amt
   Set RbCdt = GetObject("CRebateCondition", CollRebateCondition, Trim(BD.APAR_TYPE & "-" & BD.STOCK_GROUP & "-" & Level), False)
   If Not RbCdt Is Nothing Then
      P1.TOTAL_PRICE = Amt * RbCdt.REBATE_PERCENT / 100
   End If
   
   If SumTotal.Count = 0 Then
      Call SumTotal.add(P1, Trim(P1.Getkey))
   Else
      Set P2 = SumTotal(Trim(P1.Getkey))
      If P2 Is Nothing Then
         Call SumTotal.add(P1, Trim(P1.Getkey))
      Else
         P2.TOTAL_PRICE = P2.TOTAL_PRICE + P1.TOTAL_PRICE
      End If
   End If
   
   If ParentID > 0 Then
      Call Recuresive(Cm, SumTotal, GetParent(ParentID), GetCustomer(ParentID), Level + 1, BD)
   End If
End Sub
Private Function GetParent(ID As Long) As Long
Dim Cm As CNetworkChart
   Set Cm = GetObject("CNetworkChart", NetworkChartColl, Trim(Str(ID)))
   GetParent = Cm.PARENT_ID
End Function
Private Function GetCustomer(ID As Long) As Long
Dim Cm As CNetworkChart
   Set Cm = GetObject("CNetworkChart", NetworkChartColl, Trim(Str(ID)))
   GetCustomer = Cm.CUSTOMER_ID
End Function
Private Sub GenerateOrderNetworkChart(TempColl As Collection, PID As Long, Level As Long)
Dim O As CNetworkChart

   For Each O In TempColl
      If O.PARENT_ID = PID Then
         O.CUSTOMER_NAME = Space(Level * 2) & (Level + 1) & " " & O.CUSTOMER_NAME
         O.Level = Level
         Call OrderNetworkChartColl.add(O, Trim(Str(O.NETWORK_CHART_ID)))
         Call GenerateOrderNetworkChart(TempColl, O.NETWORK_CHART_ID, Level + 1)
      End If
   Next O
End Sub

