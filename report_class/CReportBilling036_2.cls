VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling036_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling036_2"
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
Private TempDateCountTagetDetails As Long
Private TempDateCountNow  As Long

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
Private m_Note As Collection

Private m_TagetDetails As Collection

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
   Set m_Note = New Collection
   
   Set m_DistinctMain = New Collection
   Set m_TagetDetails = New Collection
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
   Set m_Note = Nothing
   
   Set m_DistinctMain = Nothing
   Set m_TagetDetails = Nothing
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
   
    '------------------------------------------------------------------------------------------>          เป้าการขายเดือน
   If mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
      Call m_Headers1.AddField(Cf)
   End If
   
   If TempDateCountNow <> TempDateCountTagetDetails Then
      If mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", "มูลค่า", ">")
         Call m_Headers1.AddField(Cf)
      End If
   End If

   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))  'ตัดวงเล็บออก
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   '------------------------------------------------------------------------------------------>          ยอดขาย
    While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      If mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
         Call m_Headers1.AddField(Cf)
      End If
      If mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", "Actual", ">")
         Call m_Headers1.AddField(Cf)
      End If
   Wend
   
   If mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
   End If
   If mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "Actual", ">")
      Call m_Headers1.AddField(Cf)
   End If
   '------------------------------------------------------------------------------------------>
   TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
   TempDateString = Trim(Replace(TempDateString, ")", ""))
   
   Call Cf.SetFieldValue(6, "^", "", "<")
   Call m_Headers0.AddField(Cf)
   
   '------------------------------------------------------------------------------------------>          เป้าการขายเดือน
   If mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(4, "^", "เป้าเดือน " & IntToThaiMonth(Format(mcolParam("MONTH_ID"), "0"), 1), "<")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_AMOUNT") > 0 Or mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "เป้าเดือน " & IntToThaiMonth(Format(mcolParam("MONTH_ID"), "0"), 1), "<")
      Call m_Headers0.AddField(Cf)
   End If
   
    If TempDateCountNow <> TempDateCountTagetDetails Then
      If mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(4, "^", "เป้า " & TempDateCountNow & " วัน", "<")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_AMOUNT") > 0 Or mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", "เป้า " & TempDateCountNow & " วัน", "<")
         Call m_Headers0.AddField(Cf)
      End If
   End If
   '------------------------------------------------------------------------------------------>          ยอดขาย
   TempFromDate = mcolParam("FROM_DATE")
   While Len(TempDateString) > 0
      Call CalculateDatePeriod(TempDateString, DateCount)
      TempToDate = DateAdd("D", DateCount - 1, TempFromDate)
      TempFromDate = DateAdd("D", -DateCount + 1, TempToDate)

      If mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0 Then
         If mcolParam("PRINT_TO_FILE") = 1 Then
            Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
            Call m_Headers0.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
            Call m_Headers0.AddField(Cf)
         Else
            Call Cf.SetFieldValue(4, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
            Call m_Headers0.AddField(Cf)
         End If
      ElseIf mcolParam("SHOW_AMOUNT") > 0 Then
         Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      ElseIf mcolParam("SHOW_PRICE") > 0 Then
         Call Cf.SetFieldValue(2, "^", GetDatePeriodString(TempFromDate, TempToDate), ">")
         Call m_Headers0.AddField(Cf)
      End If
      TempFromDate = DateAdd("D", 1, TempToDate)
   Wend

   If mcolParam("SHOW_AMOUNT") > 0 And mcolParam("SHOW_PRICE") > 0 Then
      If mcolParam("PRINT_TO_FILE") = 1 Then
         Call Cf.SetFieldValue(2, "^", "รวม", ">")
         Call m_Headers0.AddField(Cf)
         Call Cf.SetFieldValue(2, "^", "รวม", ">")
         Call m_Headers0.AddField(Cf)
      Else
         Call Cf.SetFieldValue(4, "^", "รวม", ">")
         Call m_Headers0.AddField(Cf)
      End If
   ElseIf mcolParam("SHOW_AMOUNT") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   ElseIf mcolParam("SHOW_PRICE") > 0 Then
      Call Cf.SetFieldValue(2, "^", "รวม", ">")
      Call m_Headers0.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(2, "^", "ผลต่าง", ">")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "Actual", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "%", ">")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(1, "^", "Actual", ">")
   Call m_Headers1.AddField(Cf)
   
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
   ' note: SetTextCharacterExtra takes a spacing in pixels, which
    '       we get by converting from twips.
    'g_CharSpacing = -0 / Vsp.TwipsPerPixelX
   
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
   
   Call CalculateDate(mcolParam("PERIOD_DATE"), DateCount)         ' 31 วัน
   
   TempFromDate = mcolParam("FROM_DATE")
   TempToDate = DateAdd("D", DateCount - 1, mcolParam("FROM_DATE"))
   
   Dim FD As Date
   Dim LD As Date
   Dim D As String
   
   Call GetFirstLastDate(TempFromDate, FD, LD)
   TempDateCountTagetDetails = DateDiff("D", FD, LD) + 1 - Val(mcolParam("HOLIDAY"))
   TempDateCountNow = DateCount
   
   Call GetDistinctEmpAparStockCodeDocTypeFree(m_DistinctMain, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))

   Call GetSaleAmountEmpAparStockCodeDocTypeFreeDateFree(m_SaleAmounts, TempFromDate, TempToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFree(m_SaleAmountExs, TempFromDate, TempToDate, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   
   Call LoadTagetDetailEmpAparStock(m_TagetDetails, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")))
   
   If mcolParam("NOTE") > 0 Then
      Call LoadNote(m_Note, TempFromDate, TempToDate, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"))
   End If
   
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
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_BillingDoc As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim Total2(25) As Double
Dim Total3(25) As Double
Dim Total4(25) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As Long
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim TempFirstPosition As Long
Dim TempLastPosition As Long
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
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
      I = 0
      Set BodyCf = New CReportField
      I = 0
      
      For Each m_BillingDoc In m_DistinctMain
         I = I + 1

         If PrevKey3 <> m_BillingDoc.SALE_CODE And I <> 1 Then
            If mcolParam("SHOW_SALE") = 0 Then
               If Total1(TempFirstPosition) > 0 Then
                     Total1(TempLastPosition - 1) = ((Total1(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total1(TempLastPosition - 2))
                     Total1(TempLastPosition) = (Total1(TempLastPosition - 2) * 100) / ((Total1(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
               End If
               If mcolParam("SUMMARY") = 1 Then
                  Call GenerateFooter(PrevKey6 & "/" & PrevKey7, Total1, m_Details1, 0)
               Else
                  Call GenerateFooter("      รวม " & PrevKey7, Total1, m_Details1, 0)
                  Vsp.FontBold = True
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
                                                                                                                                                                                     
            k = 1
            For Each Tg In m_TagetDetails
               If Tg.Flag <> "Y" And Tg.EMPLOYEE_CODE = PrevKey3 Then
                  If mcolParam("SHOW_AMOUNT") > 0 Then
                     k = k + 1
                     Total2(k) = Total2(k) + Tg.TOTAL_AMOUNT
                     Total4(k) = Total4(k) + Tg.TOTAL_AMOUNT
                  End If
                  If mcolParam("SHOW_PRICE") > 0 Then
                     k = k + 1
                     Total2(k) = Total2(k) + Tg.TOTAL_PRICE
                     Total4(k) = Total4(k) + Tg.TOTAL_PRICE
                  End If
                  'เป้ารายวัน
                  If TempDateCountNow <> TempDateCountTagetDetails Then
                     If mcolParam("SHOW_AMOUNT") > 0 Then
                        k = k + 1
                        Total2(k) = Total2(k) + ((Tg.TOTAL_AMOUNT * TempDateCountNow) / TempDateCountTagetDetails)
                        Total4(k) = Total4(k) + ((Tg.TOTAL_AMOUNT * TempDateCountNow) / TempDateCountTagetDetails)
                     End If
                     If mcolParam("SHOW_PRICE") > 0 Then
                        k = k + 1
                        Total2(k) = Total2(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
                        Total4(k) = Total4(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
                     End If
                  End If
                  k = 1
               End If
            Next Tg
            If Total2(TempFirstPosition) > 0 Then
                  Total2(TempLastPosition - 1) = ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total2(TempLastPosition - 2))
                  Total2(TempLastPosition) = (Total2(TempLastPosition - 2) * 100) / ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
            End If
            If mcolParam("SHOW_SALE") = 1 Then
               Call GenerateFooter(PrevKey4 & " " & PrevKey3, Total2, m_Details1, 0)
            Else
               Call GenerateFooter("      รวม " & PrevKey3, Total2, m_Details1, 0)
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False

            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
            Next j
            
            If mcolParam("SHOW_SALE") = 0 Then
               Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
               If mcolParam("SUMMARY") = 0 Then
                  If mcolParam("SHORT_NAME") = 0 Then
                     Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & "/" & m_BillingDoc.APAR_CODE
                  Else
                     Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_SHORT_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & "/" & m_BillingDoc.APAR_CODE
                  End If
               End If
            End If
         ElseIf PrevKey5 <> m_BillingDoc.CUSTOMER_BRANCH And I <> 1 Then
            If mcolParam("SHOW_SALE") = 0 Then
               If Total1(TempFirstPosition) > 0 Then
                     Total1(TempLastPosition - 1) = ((Total1(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total1(TempLastPosition - 2))
                     Total1(TempLastPosition) = (Total1(TempLastPosition - 2) * 100) / ((Total1(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
               End If
               If mcolParam("SUMMARY") = 1 Then
                  Call GenerateFooter(PrevKey6 & "/" & PrevKey7, Total1, m_Details1, 0)
               Else
                  Call GenerateFooter("      รวม " & PrevKey7, Total1, m_Details1, 0)
                  Vsp.FontBold = True
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j

            If mcolParam("SUMMARY") = 0 And mcolParam("SHOW_SALE") = 0 Then
               If mcolParam("SHORT_NAME") = 0 Then
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & "/" & m_BillingDoc.APAR_CODE
               Else
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_SHORT_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & "/" & m_BillingDoc.APAR_CODE
               End If
            End If

         ElseIf I = 1 Then
            If mcolParam("PRINT_TO_FILE") Then
               Call printHeader
            End If
            If mcolParam("SHOW_SALE") = 0 Then
               Vsp = "     พนักงานขาย    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & "/" & m_BillingDoc.SALE_CODE
            End If
            If mcolParam("SUMMARY") = 0 And mcolParam("SHOW_SALE") = 0 Then
               If mcolParam("SHORT_NAME") = 0 Then
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & "/" & m_BillingDoc.APAR_CODE
               Else
                  Vsp = "           ลูกค้า        : " & m_BillingDoc.APAR_SHORT_NAME & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME & "/" & m_BillingDoc.APAR_CODE
               End If
            End If
         End If

         PrevKey3 = m_BillingDoc.SALE_CODE
         PrevKey4 = m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
         PrevKey5 = m_BillingDoc.CUSTOMER_BRANCH
         PrevKey7 = m_BillingDoc.APAR_CODE & " สาขา " & m_BillingDoc.CUSTOMER_BRANCH_NAME
         If mcolParam("SHORT_NAME") = 0 Then
            PrevKey6 = m_BillingDoc.APAR_NAME
         Else
            PrevKey6 = m_BillingDoc.APAR_SHORT_NAME
         End If

         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_BillingDoc.STOCK_DESC & "/" & m_BillingDoc.STOCK_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         k = 1
         
         'เป้าการขายเดือน
         Amt = 0
         Set Tg = GetObject("CTagetDetail", m_TagetDetails, m_BillingDoc.SALE_CODE & "-" & m_BillingDoc.APAR_CODE & "-" & m_BillingDoc.STOCK_NO)
         If mcolParam("SHOW_AMOUNT") > 0 Then
            Amt = Tg.TOTAL_AMOUNT
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            If Not Tg.Flag = "Y" Then
               TempStr = FormatNumberToNull(Amt)
            Else
               TempStr = ""
               Amt = 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
            Total3(k) = Total3(k) + Amt
            Total4(k) = Total4(k) + Amt
         End If
         If mcolParam("SHOW_PRICE") > 0 Then
            k = k + 1
            TempFirstPosition = k
            Amt = Tg.TOTAL_PRICE
            Set HeadCf = m_Headers1.Fields(k)
            If Not Tg.Flag = "Y" Then
               TempStr = FormatNumberToNull(Amt)
            Else
               TempStr = ""
               Amt = 0
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(k) = Total1(k) + Amt
            Total2(k) = Total2(k) + Amt
            Total3(k) = Total3(k) + Amt
            Total4(k) = Total4(k) + Amt
         End If
         'เป้าการขาย วัน
         If TempDateCountNow <> TempDateCountTagetDetails Then
            If mcolParam("SHOW_AMOUNT") > 0 Then
               Amt = (Tg.TOTAL_AMOUNT * TempDateCountNow) / TempDateCountTagetDetails
               k = k + 1
               Set HeadCf = m_Headers1.Fields(k)
               If Not Tg.Flag = "Y" Then
                  TempStr = FormatNumberToNull(Amt)
               Else
                  TempStr = ""
                  Amt = 0
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               Total3(k) = Total3(k) + Amt
               Total4(k) = Total4(k) + Amt
            End If
            If mcolParam("SHOW_PRICE") > 0 Then
               k = k + 1
               Amt = (Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails
               Set HeadCf = m_Headers1.Fields(k)
               If Not Tg.Flag = "Y" Then
                  TempStr = FormatNumberToNull(Amt)
               Else
                  TempStr = ""
                  Amt = 0
               End If
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(k) = Total1(k) + Amt
               Total2(k) = Total2(k) + Amt
               Total3(k) = Total3(k) + Amt
               Total4(k) = Total4(k) + Amt
            End If
         End If
         
         If Amt <> 0 Then
            Tg.Flag = "Y"
         End If
         
         TempDateString = Trim(Replace(mcolParam("PERIOD_DATE"), "(", ""))
         TempDateString = Trim(Replace(TempDateString, ")", ""))
         TempDate = DateAdd("D", -1, mcolParam("FROM_DATE"))
         SumAmount = 0
         SumTotalPrice = 0

         While Len(TempDateString) > 0
               Call CalculateDatePeriod(TempDateString, CountDate)
               TempDateCount = 0
               
                  If mcolParam("SHOW_AMOUNT") > 0 Then
                     k = k + 1
                     Amt = 0
                     CountDate1 = CountDate
                     While CountDate1 > 0
                        TempDate = DateAdd("D", 1, TempDate)
                        TempDateCount = TempDateCount + 1
                        
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_AMOUNT
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_AMOUNT
                        If m_BillingDoc.NOT_SHOW_RETURN <> "Y" Or mcolParam("NOT_SHOW_RETURN") <> "1" Then
                           Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                           Amt = Amt - TempBd.TOTAL_AMOUNT
                        End If
         
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_AMOUNT
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_AMOUNT
                        If m_BillingDoc.NOT_SHOW_RETURN <> "Y" Or mcolParam("NOT_SHOW_RETURN") <> "1" Then
                           Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                           Amt = Amt - TempBd.TOTAL_AMOUNT
                        End If
                     
                        CountDate1 = CountDate1 - 1
                     Wend
                     
                     SumAmount = SumAmount + Amt
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                     Total3(k) = Total3(k) + Amt
                     Total4(k) = Total4(k) + Amt
                     If mcolParam("SHOW_PRICE") > 0 Then
                        TempDate = DateAdd("D", -TempDateCount, TempDate)
                     End If
                  End If
                  If mcolParam("SHOW_PRICE") > 0 Then
                     k = k + 1
                     Amt = 0
                     CountDate1 = CountDate
                     While CountDate1 > 0
                        TempDate = DateAdd("D", 1, TempDate)
                        
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                        If m_BillingDoc.NOT_SHOW_RETURN <> "Y" Or mcolParam("NOT_SHOW_RETURN") <> "1" Then
                           Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                           Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
                        End If
         
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & INVOICE_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                        Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RECEIPT1_DOCTYPE & "-" & TempDate)
                        Amt = Amt + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                        If m_BillingDoc.NOT_SHOW_RETURN <> "Y" Or mcolParam("NOT_SHOW_RETURN") <> "1" Then
                           Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, GetKey1(m_BillingDoc) & "-" & RETURN_DOCTYPE & "-" & TempDate)
                           Amt = Amt - (TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT)
                        End If
                     
                        CountDate1 = CountDate1 - 1
                     Wend
                     
                     SumTotalPrice = SumTotalPrice + Amt
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                     Total2(k) = Total2(k) + Amt
                     Total3(k) = Total3(k) + Amt
                     Total4(k) = Total4(k) + Amt
                  End If
         Wend
'         '-----------------------------------------------------------------------------------------------------------------------------------------------
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
         End If
         'ผลต่าง
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         If (Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails > 0 Then
            TempStr = FormatNumberToNull(SumTotalPrice - ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails))
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         '%
         k = k + 1
         TempLastPosition = k
         Set HeadCf = m_Headers1.Fields(k)
         If (Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails > 0 Then
            TempStr = FormatNumberToNull((SumTotalPrice * 100) / ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails))
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + 0
         Total2(k) = Total2(k) + 0
         Total3(k) = Total3(k) + 0
         Total4(k) = Total4(k) + 0

         If mcolParam("SUMMARY") = 0 And mcolParam("SHOW_SALE") = 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
      Next m_BillingDoc
      
      If mcolParam("SHOW_SALE") = 0 Then
         If Total1(TempFirstPosition) > 0 Then
            Total1(TempLastPosition - 1) = ((Total1(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total1(TempLastPosition - 2))
            Total1(TempLastPosition) = (Total1(TempLastPosition - 2) * 100) / ((Total1(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
         End If
         If mcolParam("SUMMARY") = 1 Then
            Call GenerateFooter(PrevKey6 & "/" & PrevKey7, Total1, m_Details1, 0)
         Else
            Call GenerateFooter("      รวม " & PrevKey7, Total1, m_Details1, 0)
            Vsp.FontBold = True
         End If
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
      k = 1
      For Each Tg In m_TagetDetails
         If Tg.Flag <> "Y" And Tg.EMPLOYEE_CODE = PrevKey3 Then
            If mcolParam("SHOW_AMOUNT") > 0 Then
               k = k + 1
               Total2(k) = Total2(k) + Tg.TOTAL_AMOUNT
               Total4(k) = Total4(k) + Tg.TOTAL_AMOUNT
            End If
            If mcolParam("SHOW_PRICE") > 0 Then
               k = k + 1
               Total2(k) = Total2(k) + Tg.TOTAL_PRICE
               Total4(k) = Total4(k) + Tg.TOTAL_PRICE
            End If
            'เป้ารายวัน
            If TempDateCountNow <> TempDateCountTagetDetails Then
               If mcolParam("SHOW_AMOUNT") > 0 Then
                  k = k + 1
                  Total2(k) = Total2(k) + ((Tg.TOTAL_AMOUNT * TempDateCountNow) / TempDateCountTagetDetails)
                  Total4(k) = Total4(k) + ((Tg.TOTAL_AMOUNT * TempDateCountNow) / TempDateCountTagetDetails)
               End If
               If mcolParam("SHOW_PRICE") > 0 Then
                  k = k + 1
                  Total2(k) = Total2(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
                  Total4(k) = Total4(k) + ((Tg.TOTAL_PRICE * TempDateCountNow) / TempDateCountTagetDetails)
               End If
            End If
            k = 1
         End If
      Next Tg
      
      If Total2(TempFirstPosition) > 0 Then
         Total2(TempLastPosition - 1) = ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total2(TempLastPosition - 2))
         Total2(TempLastPosition) = (Total2(TempLastPosition - 2) * 100) / ((Total2(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
      End If
      If mcolParam("SHOW_SALE") = 1 Then
         Call GenerateFooter(PrevKey4 & " " & PrevKey3, Total2, m_Details1, 0)
      Else
         Call GenerateFooter("      รวม " & PrevKey3, Total2, m_Details1, 0)
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      If Total4(TempFirstPosition) > 0 Then
         Total4(TempLastPosition - 1) = ((Total4(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails - Total4(TempLastPosition - 2))
         Total4(TempLastPosition) = (Total4(TempLastPosition - 2) * 100) / ((Total4(TempFirstPosition) * TempDateCountNow) / TempDateCountTagetDetails)
      End If
      Call GenerateFooter("      สุทธิ" & "" & "", Total4, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      PrevKey1 = ""
      If mcolParam("NOTE") > 0 Then
         Vsp = "หมายเหตุ"
         For Each m_BillingDoc In m_Note
               PrevKey1 = PrevKey1 & m_BillingDoc.DOCUMENT_NO & " / " & DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE) & " / " & m_BillingDoc.NOTE & "        "
         Next m_BillingDoc
         Vsp = PrevKey1
      End If

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
   GetKey1 = Trim(Tg.SALE_CODE & "-" & Tg.APAR_CODE & "-" & Tg.CUSTOMER_BRANCH & "-" & Tg.STOCK_NO)
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
