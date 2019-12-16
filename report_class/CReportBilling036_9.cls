VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling036_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling036_9"
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

Private m_DistinctMain As Collection
Private m_SaleAmounts As Collection
Private m_SaleAmountExs As Collection
Private OtherMoves As Collection
Private m_TagetDetails As Collection
Private NotStockNo As Collection
Private DateCount As Long
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
   
   Set m_SaleAmounts = New Collection
   Set m_SaleAmountExs = New Collection
   Set OtherMoves = New Collection
   Set m_DistinctMain = New Collection
   Set m_TagetDetails = New Collection
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
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_SaleAmounts = Nothing
   Set m_SaleAmountExs = Nothing
   Set OtherMoves = Nothing
   Set m_DistinctMain = Nothing
   Set m_TagetDetails = Nothing
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
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long
Dim DateCount  As Long
Dim TempDateString As String
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
   Call m_Headers1_1.ClearField
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(6, "^", "รายละเอียด", "<")
   Call m_Headers1.AddField(Cf)
   
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   '------------------------------------------------------------------------------------------>          ยอดขาย
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      If mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("SHOW_COST") > 0 Then
         Call Cf.SetFieldValue(2, "^", "ต้นทุน", ">")
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("SHOW_GP") > 0 Then
         Call Cf.SetFieldValue(2, "^", "กำไร", ">")
         Call m_Headers1.AddField(Cf)
      End If
   Wend
   If mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_COST") > 0 Then
      Call Cf.SetFieldValue(2, "^", "ต้นทุน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_GP") > 0 Then
      Call Cf.SetFieldValue(2, "^", "กำไร", ">")
      Call m_Headers1.AddField(Cf)
   End If
   '------------------------------------------------------------------------------------------>
   
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   Call Cf.SetFieldValue(6, "^", "", "<")
   Call m_Headers0.AddField(Cf)
   '------------------------------------------------------------------------------------------>          ยอดขาย
   TempFromDate = mcolParam("FROM_BILL_DATE")
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
      TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)
      
      If mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(8, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0) Or (mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0) Then
         Call Cf.SetFieldValue(6, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_AMOUNT") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0) Or (mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0) Then
         Call Cf.SetFieldValue(4, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_COST") > 0 Then
         Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_GP") > 0 Then
         Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      End If
      TempFromDate = DateAdd("D", 1, TempToDate)
   Wend
   If mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(8, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0) Or (mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0) Then
      Call Cf.SetFieldValue(6, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_COST") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_AMOUNT") > 0) Or (mcolParam("SHOW_GP") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_AMOUNT") > 0) Or (mcolParam("SHOW_COST") > 0 And mcolParam("SHOW_PRICE") > 0) Or (mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0) Then
      Call Cf.SetFieldValue(4, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_COST") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_GP") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
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
Dim TempFromDate  As Date
Dim TempToDate As Date
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

   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)
   
   TempFromDate = mcolParam("FROM_BILL_DATE")
   TempToDate = DateAdd("D", DateCount - 1, mcolParam("FROM_BILL_DATE"))

   Call LoadStockCodeFromTo(Nothing, NotStockNo, mcolParam("NOT_FROM_STOCK_NO"), mcolParam("NOT_TO_STOCK_NO"), 1)

   Call GetDistinctEmpAparStockCodeDocTypeFreeNotBranch(m_DistinctMain, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))

   Call GetSaleAmountEmpAparStockCodeDocTypeFreeDateFreeNotBranch(m_SaleAmounts, TempFromDate, TempToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFreeNotBranch(m_SaleAmountExs, TempFromDate, TempToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   
   Call GetSaleAmountStockCodeDocTypeDate(OtherMoves, TempFromDate, TempToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   
   FromDate = "จากงวด " & EmptyToString(DateToStringExtEx2(TempFromDate), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(TempToDate), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " จากพนักงานขาย " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " ถึงพนักงานขาย " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
                                       
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
Dim TempCheck As Long
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
Dim Total3(100) As Double
Dim Total4(100) As Double
Dim Total5(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim PrevKey8 As String
Dim CheckStock As Boolean
Dim TempBillingDoc As CBillingDoc
Dim TempBillingDoc2 As CBillingDoc
Dim TempBillingDoc3 As CBillingDoc
Dim SumItem As Double
Dim CostAmount As Double
Dim TempAmount As Double

Dim SumCost As Double
Dim SumAmount As Double
Dim SumTotalPrice As Double

Dim TempBd As CBillingDoc

Dim AmountDate As Long
Dim Ratio As Double

Dim TempDateString As String
Dim CountDate As Long
Dim CountDate1 As Long
Dim TempDate As Date
Dim TempDateCount As Long
Dim Tg As CTagetDetail
   
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
      Total5(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField
      I = 0
      If mcolParam("SHOW_GROUP_CUSTOMER") = 1 And mcolParam("SUMMARY") = 1 Then
         TempCheck = 1
      Else
         TempCheck = 0
      End If
      
      For Each m_BillingDoc In m_DistinctMain
         I = I + 1

         If PrevKey3 <> m_BillingDoc.SALE_CODE And I <> 1 Then

            If mcolParam("SUMMARY") = 1 Then
               If mcolParam("SHOW_GROUP_CUSTOMER") = 1 Then
                  If PrevKey8 = "" Then
                     Call GenerateFooter(PrevKey6, Total5, m_Details1, 0)
                  Else
                     Call GenerateFooter(PrevKey8, Total5, m_Details1, 0)
                  End If
               Else
                  Call GenerateFooter(PrevKey6 & "/" & PrevKey5, Total1, m_Details1, 0)
               End If
            Else
               Call GenerateFooter("      รวม" & PrevKey5, Total1, m_Details1, 0)
               Vsp.FontBold = True
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            Call GenerateFooter("      รวม" & PrevKey3, Total2, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
               Total5(j) = 0
            Next j

            Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            If mcolParam("SUMMARY") = 0 Then
               If mcolParam("SHORT_NAME") = 0 Then
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_NAME & "/" & m_BillingDoc.APAR_CODE
               Else
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_SHORT_NAME & "/" & m_BillingDoc.APAR_CODE
               End If
            End If
         ElseIf mcolParam("SHOW_GROUP_CUSTOMER") = 1 And mcolParam("SUMMARY") = 1 And PrevKey7 <> m_BillingDoc.APAR_MAS_GROUP_CODE And I <> 1 Then
               If PrevKey8 = "" Then
                  Call GenerateFooter(PrevKey6, Total5, m_Details1, 0)
               Else
                  Call GenerateFooter(PrevKey8, Total5, m_Details1, 0)
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               For j = 1 To UBound(Total5)
                  Total5(j) = 0
               Next j
         ElseIf PrevKey5 <> m_BillingDoc.APAR_CODE And I <> 1 And TempCheck = 0 Then
               If mcolParam("SUMMARY") = 1 Then
                  Call GenerateFooter(PrevKey6 & "/" & PrevKey5, Total1, m_Details1, 0)
               Else
                  Call GenerateFooter("      รวม" & PrevKey5, Total1, m_Details1, 0)
                  Vsp.FontBold = True
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j

            If mcolParam("SUMMARY") = 0 Then
               If mcolParam("SHORT_NAME") = 0 Then
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_NAME & "/" & m_BillingDoc.APAR_CODE
               Else
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_SHORT_NAME & "/" & m_BillingDoc.APAR_CODE
               End If
            End If

         ElseIf I = 1 Then
            If mcolParam("PRINT_TO_EXCEL") = 1 Then
               Call printHeader
            End If
            Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            If mcolParam("SUMMARY") = 0 Then
               If mcolParam("SHORT_NAME") = 0 Then
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_NAME & "/" & m_BillingDoc.APAR_CODE
               Else
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_SHORT_NAME & "/" & m_BillingDoc.APAR_CODE
               End If
            End If
         End If

         PrevKey3 = m_BillingDoc.SALE_CODE
         PrevKey4 = m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
         PrevKey5 = m_BillingDoc.APAR_CODE
         If mcolParam("SHORT_NAME") = 0 Then
            PrevKey6 = m_BillingDoc.APAR_NAME
         Else
            PrevKey6 = m_BillingDoc.APAR_SHORT_NAME
         End If
         PrevKey7 = m_BillingDoc.APAR_MAS_GROUP_CODE
         PrevKey8 = m_BillingDoc.APAR_MAS_GROUP_NAME
         Call m_Details1.ClearField

         Call CheckNotStock(m_BillingDoc.STOCK_NO, CheckStock)
         If CheckStock = True Then
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = m_BillingDoc.STOCK_DESC & "/" & m_BillingDoc.STOCK_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            k = 1
   
            TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
            TempDateString = Trim(Replace(TempDateString, ")", ""))
            TempDate = DateAdd("D", -1, mcolParam("FROM_BILL_DATE"))
            SumAmount = 0
            SumTotalPrice = 0
            SumCost = 0
            
            While Len(TempDateString) > 0
                  Call CalculateDatePeriod(TempDateString, CountDate)
                  TempDateCount = 0
                  Amt = 0
                  TempAmount = 0
                  CountDate1 = CountDate
                  While CountDate1 > 0
                     TempDate = DateAdd("D", 1, TempDate)
                     TempDateCount = TempDateCount + 1
                           
                     Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                     Amt = Amt + TempBd.TOTAL_AMOUNT
                     Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                     Amt = Amt + TempBd.TOTAL_AMOUNT
                     Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                     Amt = Amt - TempBd.TOTAL_AMOUNT
            
                     Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                     Amt = Amt + TempBd.TOTAL_AMOUNT
                     Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                     Amt = Amt + TempBd.TOTAL_AMOUNT
                     Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                     Amt = Amt - TempBd.TOTAL_AMOUNT
                     CountDate1 = CountDate1 - 1
                  Wend
                  
                  SumAmount = SumAmount + Amt
                  TempAmount = Amt
                  TempDate = DateAdd("D", -TempDateCount, TempDate)
                  If mcolParam("SHOW_AMOUNT") > 0 Then
                     k = k + 1
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                     Total3(k) = Total3(k) + Amt
                     Total4(k) = Total4(k) + Amt
                     Total5(k) = Total5(k) + Amt
                  End If
            
               Amt = 0
               CountDate1 = CountDate
               While CountDate1 > 0
                  TempDate = DateAdd("D", 1, TempDate)
                  Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                  Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
            
                  Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                  Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                  Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
                  CountDate1 = CountDate1 - 1
               Wend
               
               SumTotalPrice = SumTotalPrice + Amt
               If mcolParam("SHOW_COST") > 0 Then
                  TempDate = DateAdd("D", -TempDateCount, TempDate)
               End If
               If mcolParam("SHOW_PRICE") > 0 Then
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(Amt)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + Amt
                  Total2(k) = Total2(k) + Amt
                  Total3(k) = Total3(k) + Amt
                  Total4(k) = Total4(k) + Amt
                  Total5(k) = Total5(k) + Amt
               End If
   
               If mcolParam("SHOW_COST") > 0 Then
                  k = k + 1
                  CostAmount = 0
                  SumItem = 0
   
                  CountDate1 = CountDate
                  While CountDate1 > 0
                     TempDate = DateAdd("D", 1, TempDate)
                     Set TempBillingDoc = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & INVOICE_DOCTYPE & "-" & TempDate), True)
                     Set TempBillingDoc2 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RECEIPT1_DOCTYPE & "-" & TempDate), True)
                     Set TempBillingDoc3 = GetObject("CBillingDoc", OtherMoves, Trim(m_BillingDoc.STOCK_NO & "-" & RETURN_DOCTYPE & "-" & TempDate), True)
                     
                     SumItem = SumItem + TempBillingDoc.TOTAL_AMOUNT + TempBillingDoc2.TOTAL_AMOUNT - TempBillingDoc3.TOTAL_AMOUNT
                      
                     CostAmount = CostAmount + (TempBillingDoc.TOTAL_INCLUDE_PRICE) + (TempBillingDoc2.TOTAL_INCLUDE_PRICE) - (TempBillingDoc3.TOTAL_INCLUDE_PRICE)
                     CountDate1 = CountDate1 - 1
                  Wend
                  
                  SumCost = SumCost + (TempAmount * MyDiffEx(CostAmount, SumItem))
                  
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(TempAmount * MyDiffEx(CostAmount, SumItem))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + (TempAmount * MyDiffEx(CostAmount, SumItem))
                  Total2(k) = Total2(k) + (TempAmount * MyDiffEx(CostAmount, SumItem))
                  Total3(k) = Total3(k) + (TempAmount * MyDiffEx(CostAmount, SumItem))
                  Total4(k) = Total4(k) + (TempAmount * MyDiffEx(CostAmount, SumItem))
                  Total5(k) = Total5(k) + (TempAmount * MyDiffEx(CostAmount, SumItem))
               End If
               If mcolParam("SHOW_GP") > 0 Then
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(Amt - (TempAmount * MyDiffEx(CostAmount, SumItem)))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + (Amt - (TempAmount * MyDiffEx(CostAmount, SumItem)))
                  Total2(k) = Total2(k) + (Amt - (TempAmount * MyDiffEx(CostAmount, SumItem)))
                  Total3(k) = Total3(k) + (Amt - (TempAmount * MyDiffEx(CostAmount, SumItem)))
                  Total4(k) = Total4(k) + (Amt - (TempAmount * MyDiffEx(CostAmount, SumItem)))
                  Total5(k) = Total5(k) + (Amt - (TempAmount * MyDiffEx(CostAmount, SumItem)))
               End If
            Wend
            '-----------------------------------------------------------------------------------------------------------------------------------------------
            If mcolParam("SHOW_AMOUNT") > 0 Then
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(SumAmount)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + SumAmount
               Total2(k) = Total2(k) + SumAmount
               Total3(k) = Total3(k) + SumAmount
               Total4(k) = Total4(k) + SumAmount
               Total5(k) = Total5(k) + SumAmount
            End If
            If mcolParam("SHOW_PRICE") > 0 Then
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(SumTotalPrice)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + SumTotalPrice
               Total2(k) = Total2(k) + SumTotalPrice
               Total3(k) = Total3(k) + SumTotalPrice
               Total4(k) = Total4(k) + SumTotalPrice
               Total5(k) = Total5(k) + SumTotalPrice
            End If
            If mcolParam("SHOW_COST") > 0 Then
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(SumCost)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + SumCost
               Total2(k) = Total2(k) + SumCost
               Total3(k) = Total3(k) + SumCost
               Total4(k) = Total4(k) + SumCost
               Total5(k) = Total5(k) + SumCost
            End If
            If mcolParam("SHOW_GP") > 0 Then
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               TempStr = FormatNumberToNull(SumTotalPrice - SumCost)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + SumTotalPrice - SumCost
               Total2(k) = Total2(k) + SumTotalPrice - SumCost
               Total3(k) = Total3(k) + SumTotalPrice - SumCost
               Total4(k) = Total4(k) + SumTotalPrice - SumCost
               Total5(k) = Total5(k) + SumTotalPrice - SumCost
            End If
            If mcolParam("SUMMARY") = 0 Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
      Next m_BillingDoc

      If mcolParam("SUMMARY") = 1 Then
         If mcolParam("SHOW_GROUP_CUSTOMER") = 1 Then
            If PrevKey8 = "" Then
               Call GenerateFooter(PrevKey6, Total5, m_Details1, 0)
            Else
               Call GenerateFooter(PrevKey8, Total5, m_Details1, 0)
            End If
         Else
            Call GenerateFooter(PrevKey6 & "/" & PrevKey5, Total1, m_Details1, 0)
         End If
      Else
         Call GenerateFooter("      รวม" & PrevKey5, Total1, m_Details1, 0)
         Vsp.FontBold = True
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("      รวม" & PrevKey3, Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False

      Call GenerateFooter("      สุทธิ" & "" & "", Total4, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateEndDoc
      
      Set m_SaleAmounts = Nothing
      Set m_SaleAmountExs = Nothing
      
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
      ElseIf j >= 2 Then
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
Private Function GetKey1(Tg As CBillingDoc) As String
   GetKey1 = Trim(Tg.SALE_CODE & "-" & Tg.APAR_CODE & "-" & Tg.STOCK_NO)
End Function
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
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
