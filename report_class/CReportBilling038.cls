VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling038"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling038"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<52|<3|<5;"

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

Private m_SaleAmounts As Collection
Private m_SaleAmountExs As Collection

Private m_SaleAmountCompares As Collection
Private m_SaleAmountExCompares As Collection

Private m_SaleAparAmounts As Collection
Private m_SaleAparAmountExs As Collection

Private m_SaleAparAmountCompares As Collection
Private m_SaleAparAmountExCompares As Collection

Private m_Taget As Collection
Private m_TagetTop As Collection
Private CanPrintHeaderKeyAccount As Boolean

' note: this API is declared incorrectly in the VB API Viewer.
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   
   Set m_SaleAmounts = New Collection
   Set m_SaleAmountExs = New Collection
   
   Set m_SaleAmountCompares = New Collection
   Set m_SaleAmountExCompares = New Collection
   
   Set m_SaleAparAmounts = New Collection
   Set m_SaleAparAmountExs = New Collection
   
   Set m_SaleAparAmountCompares = New Collection
   Set m_SaleAparAmountExCompares = New Collection
   
   Set m_Taget = New Collection
   Set m_TagetTop = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   
   Set m_SaleAmounts = Nothing
   Set m_SaleAmountExs = Nothing
   Set m_SaleAmountCompares = Nothing
   Set m_SaleAmountExCompares = Nothing
   
   Set m_SaleAparAmounts = Nothing
   Set m_SaleAparAmountExs = Nothing
   
   Set m_SaleAparAmountCompares = Nothing
   Set m_SaleAparAmountExCompares = Nothing
   
   Set m_Taget = Nothing
   Set m_TagetTop = Nothing
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
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim ShowFlagAmount As Long
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(5, "^", "สาขา", "<")
   Call m_Headers1.AddField(Cf)
   
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
      
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   Call Cf.SetFieldValue(5, "^", "", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", GetDatePeriodString(mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE")), "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", GetDatePeriodString(mcolParam("FROM_BILL_DATE_EX"), mcolParam("TO_BILL_DATE_EX")), "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", "ผลต่าง", "<")
   Call m_Headers0.AddField(Cf)
   
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
Private Sub printHeader1()
Dim strFormat As String
Dim strPrint As String
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim TempStr As String
Dim ShowFlagAmount As Long
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers1_1.ClearField
   Call m_Headers1_2.ClearField
   
   Set Cf = New CReportField
   
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
      
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   
   
   
   ShowFlagAmount = 0
   If mcolParam("SHOW_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จำนวน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดเงิน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
      Call Cf.SetFieldValue(2, "^", "จ.คืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   If mcolParam("SHOW_RETURN_PRICE") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ยอดคืน", ">")
      Call m_Headers1_2.AddField(Cf)
      ShowFlagAmount = ShowFlagAmount + 1
   End If
   
   
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", "เป้าการขาย  " & IntToThaiMonth(mcolParam("MONTH_ID")) & "/" & mcolParam("YEAR_NO"), "<")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", GetDatePeriodString(mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE")), "<")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", GetDatePeriodString(mcolParam("FROM_BILL_DATE_EX"), mcolParam("TO_BILL_DATE_EX")), "<")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(2 * ShowFlagAmount, "^", "%เปรียบเทียบ", "<")
   Call m_Headers1_1.AddField(Cf)
   
   Call m_Headers1_1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers1_2.GetString(1, TempStr1, TempStr1_1)
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

Dim FromDate As String
Dim ToDate As String
Dim TempFromDate  As String
Dim TempToDate As String
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
   
   Call GetSaleAmountEmpDocType(m_SaleAmounts, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountEmpDocTypeBranch(m_SaleAmountExs, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   
   Call GetSaleAmountEmpDocType(m_SaleAmountCompares, mcolParam("FROM_BILL_DATE_EX"), mcolParam("TO_BILL_DATE_EX"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountEmpDocTypeBranch(m_SaleAmountExCompares, mcolParam("FROM_BILL_DATE_EX"), mcolParam("TO_BILL_DATE_EX"), mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
      
   Call GetSaleAmountEmpAparBranchCode(m_SaleAparAmounts, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), "", "", mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountEmpAparBranchCodeBranch(m_SaleAparAmountExs, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), "", "", mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   
   Call GetSaleAmountEmpAparBranchCode(m_SaleAparAmountCompares, mcolParam("FROM_BILL_DATE_EX"), mcolParam("TO_BILL_DATE_EX"), "", "", mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountEmpAparBranchCodeBranch(m_SaleAparAmountExCompares, mcolParam("FROM_BILL_DATE_EX"), mcolParam("TO_BILL_DATE_EX"), "", "", mcolParam("FROM_SALE_CODE"), mcolParam("TO_SALE_CODE"), mcolParam("INCLUDE_FREE"))
   
   Call LoadTaget(m_Taget, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")))
   
   Call LoadTaget(m_TagetTop, Trim((Val(mcolParam("YEAR_NO")) - 543) & Format(mcolParam("MONTH_ID"), "00")), "Y")
   
   FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   TempFromDate = "จากวันที่เปรียบเทียบ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE_EX")), "N/A")
   TempToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE_EX")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " " & TempFromDate & " " & TempToDate & "  เป้าการขาย  " & IntToThaiMonth(mcolParam("MONTH_ID")) & "/" & mcolParam("YEAR_NO")
   
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(25) As Double
Dim Total2(25) As Double
Dim iCount As Long
Dim TempStr As String

'---------------------------------------------------------
Dim Emp As CEmployee
Dim TempBd As CBillingDoc

Dim SumAmount As Double
Dim SumAmountCompare As Double
Dim SumAmountTop As Double
Dim SumAmountTopCompare As Double

Dim SumPrice As Double
Dim SumPriceTop As Double
Dim SumPriceCompare As Double
Dim SumPriceTopCompare As Double

Dim SumReturnAmount As Double
Dim SumReturnAmountTop As Double
Dim SumReturnAmountComPare As Double
Dim SumReturnAmountTopComPare As Double

Dim SumReturnPrice As Double
Dim SumReturnPriceCompare As Double
Dim SumReturnPriceTop As Double
Dim SumReturnPriceTopCompare As Double

Dim X As Long
Dim Y As Long
Dim m_KeyAccount As CKeyAccount
Dim m_KeyAccountDetail As CKeyAccountDetail
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long

Dim Tgd As CTagetDetail
'---------------------------------------------------------

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
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
      
      I = 0
      
      For Each Emp In m_EmployeeColl
         If CheckInSaleCode(Emp.EMP_CODE) Then
            '------------------------------------  เซ็ตค่าต่างเป็น 0 --------------------------------------------------------'
            SumAmount = 0
            SumAmountCompare = 0
            SumAmountTop = 0
            SumAmountTopCompare = 0
            
            SumPrice = 0
            SumPriceCompare = 0
            SumPriceTop = 0
            SumPriceTopCompare = 0
            
            SumReturnAmount = 0
            SumReturnAmountComPare = 0
            SumReturnAmountTop = 0
            SumReturnAmountTopComPare = 0
            
            SumReturnPrice = 0
            SumReturnPriceCompare = 0
            SumReturnPriceTop = 0
            SumReturnPriceTopCompare = 0
            '------------------------------------  เซ็ตค่าต่างเป็น 0 --------------------------------------------------------'
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------------------------------'
            Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Emp.EMP_CODE & "-" & INVOICE_DOCTYPE)
            SumAmount = SumAmount + TempBd.TOTAL_AMOUNT
            SumPrice = SumPrice + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                     
            Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Emp.EMP_CODE & "-" & RECEIPT1_DOCTYPE)
            SumAmount = SumAmount + TempBd.TOTAL_AMOUNT
            SumPrice = SumPrice + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            
            Set TempBd = GetObject("CBillingDoc", m_SaleAmounts, Emp.EMP_CODE & "-" & RETURN_DOCTYPE)
            SumReturnAmount = SumReturnAmount + TempBd.TOTAL_AMOUNT
            SumReturnPrice = SumReturnPrice + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------------------------------'
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----------ของ TOP--------------------'
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, Emp.EMP_CODE & "-" & INVOICE_DOCTYPE)
            SumAmountTop = SumAmountTop + TempBd.TOTAL_AMOUNT
            SumPriceTop = SumPriceTop + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                     
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, Emp.EMP_CODE & "-" & RECEIPT1_DOCTYPE)
            SumAmountTop = SumAmountTop + TempBd.TOTAL_AMOUNT
            SumPriceTop = SumPriceTop + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExs, Emp.EMP_CODE & "-" & RETURN_DOCTYPE)
            SumReturnAmountTop = SumReturnAmountTop + TempBd.TOTAL_AMOUNT
            SumReturnPriceTop = SumReturnPriceTop + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----------ของ TOP--------------------'
            
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน---------------เปรียบเทียบ----------------'
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountCompares, Emp.EMP_CODE & "-" & INVOICE_DOCTYPE)
            SumAmountCompare = SumAmountCompare + TempBd.TOTAL_AMOUNT
            SumPriceCompare = SumPriceCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                     
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountCompares, Emp.EMP_CODE & "-" & RECEIPT1_DOCTYPE)
            SumAmountCompare = SumAmountCompare + TempBd.TOTAL_AMOUNT
            SumPriceCompare = SumPriceCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountCompares, Emp.EMP_CODE & "-" & RETURN_DOCTYPE)
            SumReturnAmountComPare = SumReturnAmountComPare + TempBd.TOTAL_AMOUNT
            SumReturnPriceCompare = SumReturnPriceCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน---------------เปรียบเทียบ----------------'
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----------ของ TOP---เปรียบเทียบ-----------------'
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExCompares, Emp.EMP_CODE & "-" & INVOICE_DOCTYPE)
            SumAmountTopCompare = SumAmountTopCompare + TempBd.TOTAL_AMOUNT
            SumPriceTopCompare = SumPriceTopCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                     
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExCompares, Emp.EMP_CODE & "-" & RECEIPT1_DOCTYPE)
            SumAmountTopCompare = SumAmountTopCompare + TempBd.TOTAL_AMOUNT
            SumPriceTopCompare = SumPriceTopCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            
            Set TempBd = GetObject("CBillingDoc", m_SaleAmountExCompares, Emp.EMP_CODE & "-" & RETURN_DOCTYPE)
            SumReturnAmountTopComPare = SumReturnAmountTopComPare + TempBd.TOTAL_AMOUNT
            SumReturnPriceTopCompare = SumReturnPriceTopCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
            '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----------ของ TOP---เปรียบเทียบ-----------------'
            
            If (SumAmount + SumAmountTop + SumPrice + SumPriceTop + SumAmountCompare + SumAmountTopCompare + SumPriceCompare + SumPriceTopCompare + _
            SumReturnAmount + SumReturnAmount + SumReturnPrice + SumReturnPriceTop + SumReturnAmountComPare + SumReturnAmountComPare + SumReturnPriceCompare + SumReturnPriceTopCompare) <> 0 Then
               I = I + 1
               
               If I > 1 Then
                  CanPrintHeaderKeyAccount = False
                  Vsp.NewPage
                  Vsp.TextAlign = taCenterMiddle
                  Vsp = "ผู้รับผิดชอบ       " & Emp.EMP_NAME
                  Vsp.TextAlign = taLeftMiddle
                  Call printHeader1
               Else
                  CanPrintHeaderKeyAccount = False
                  Vsp.TextAlign = taCenterMiddle
                  Vsp = "ผู้รับผิดชอบ       " & Emp.EMP_NAME
                  Vsp.TextAlign = taLeftMiddle
                  Call printHeader1
               End If
                              
               strFormat = VSP_CalTable("10;", mdblWid, alngX)
               TempStr = "ยอดขายรวม;"
               Call Vsp.AddTable(strFormat, "", TempStr)
               
               'ตารางยอดขาย
               '------------------------------------------------------------------------------------------------------------------------
               Call m_Details1.ClearField
                  
               Set Tgd = GetObject("CTagetDetail", m_Taget, Trim(Str(Emp.EMP_ID)))
               k = 0
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_AMOUNT"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_PRICE"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_AMOUNT_RT"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_PRICE_RT"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumAmount + SumAmountTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumPrice + SumPriceTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
              End If
              
              If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnAmount + SumReturnAmountTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnPrice + SumReturnPriceTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumAmountCompare + SumAmountTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumPriceCompare + SumPriceTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnAmountComPare + SumReturnAmountTopComPare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnPriceCompare + SumReturnPriceTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumAmount + SumAmountTop, SumAmountCompare + SumAmountTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumPrice + SumPriceTop, SumPriceCompare + SumPriceTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumReturnAmount + SumReturnAmountTop, SumReturnAmountComPare + SumReturnAmountTopComPare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumReturnAmount + SumReturnAmountTop, SumReturnAmountComPare + SumReturnAmountTopComPare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               '------------------------------------------------------------------------------------------------------------------------
               
               ' % เป้าการขาย
               Call m_Details1.ClearField
               k = 0
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumAmount + SumAmountTop) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumPrice + SumPriceTop) * 100, Tgd.GetFieldValue("TOTAL_PRICE")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
              End If
              
              If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnAmount + SumReturnAmountTop) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnPrice + SumReturnPriceTop) * 100, Tgd.GetFieldValue("TOTAL_PRICE_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumAmountCompare + SumAmountTopCompare) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumPriceCompare + SumPriceTopCompare) * 100, Tgd.GetFieldValue("TOTAL_PRICE")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnAmountComPare + SumReturnAmountTopComPare) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnPriceCompare + SumReturnPriceTopCompare) * 100, Tgd.GetFieldValue("TOTAL_PRICE_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               '------------------------------------------------------------------------------------------------------------------------
               
               strFormat = VSP_CalTable("10;", mdblWid, alngX)
               TempStr = "ยอดขาย TOPS;"
               Call Vsp.AddTable(strFormat, "", TempStr)
               
               Set Tgd = GetObject("CTagetDetail", m_TagetTop, Trim(Str(Emp.EMP_ID)))
               'ตารางยอดขาย   TOPS
               '------------------------------------------------------------------------------------------------------------------------
               Call m_Details1.ClearField
               
               k = 0
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_AMOUNT"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_PRICE"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_AMOUNT_RT"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(Tgd.GetFieldValue("TOTAL_PRICE_RT"))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumAmountTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumPriceTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
              End If
              
              If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnAmountTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnPriceTop)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumAmountTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumPriceTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnAmountTopComPare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(SumReturnPriceTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumAmountTop, SumAmountTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumPriceTop, SumPriceTopCompare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumReturnAmountTop, SumReturnAmountTopComPare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = GenerateCompareString("", SumReturnAmountTop, SumReturnAmountTopComPare)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               '------------------------------------------------------------------------------------------------------------------------
               
               '
               '------------------------------------------------------------------------------------------------------------------------
               
               ' % เป้าการขาย TOP
               Call m_Details1.ClearField
               k = 0
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumAmountTop) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumPriceTop) * 100, Tgd.GetFieldValue("TOTAL_PRICE")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
              End If
              
              If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnAmountTop) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnPriceTop) * 100, Tgd.GetFieldValue("TOTAL_PRICE_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumAmountTopCompare) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumPriceTopCompare) * 100, Tgd.GetFieldValue("TOTAL_PRICE")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnAmountTopComPare) * 100, Tgd.GetFieldValue("TOTAL_AMOUNT_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((SumReturnPriceTopCompare) * 100, Tgd.GetFieldValue("TOTAL_PRICE_RT")))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                  k = k + 1
                  Set HeadCf = m_Headers1_2.Fields(k)
                  TempStr = ""
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               '------------------------------------------------------------------------------------------------------------------------
               
               '
               
               Vsp = "หมาย เหตุ     :"
               Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
               Vsp = ""
               Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
               
               Set m_KeyAccount = New CKeyAccount
               Set m_Rs = New ADODB.Recordset
               m_KeyAccount.QueryFlag = 1
               m_KeyAccount.SALE_ID = Emp.EMP_ID
               Call glbDaily.QueryKeyAccount(m_KeyAccount, m_Rs, ItemCount, IsOK, glbErrorLog)
               
               If m_KeyAccount.KeyAccountDetail.Count > 0 Then
                  Vsp = ""
                  Vsp = "สรุป KEY ACCOUNT"
                  Call printHeader
                  CanPrintHeaderKeyAccount = True
               End If
               
               For j = 1 To UBound(Total2)
                  Total1(j) = 0
               Next j
               
               For Each m_KeyAccountDetail In m_KeyAccount.KeyAccountDetail
                  '------------------------------------  เซ็ตค่าต่างเป็น 0 --------------------------------------------------------'
                  SumAmount = 0
                  SumAmountCompare = 0
                  SumAmountTop = 0
                  SumAmountTopCompare = 0
                  
                  SumPrice = 0
                  SumPriceCompare = 0
                  SumPriceTop = 0
                  SumPriceTopCompare = 0
                  
                  SumReturnAmount = 0
                  SumReturnAmountComPare = 0
                  SumReturnAmountTop = 0
                  SumReturnAmountTopComPare = 0
                  
                  SumReturnPrice = 0
                  SumReturnPriceCompare = 0
                  SumReturnPriceTop = 0
                  SumReturnPriceTopCompare = 0
                  '------------------------------------  เซ็ตค่าต่างเป็น 0 --------------------------------------------------------'
            
                  Call m_Details1.ClearField
                  
                  Set HeadCf = m_Headers1.Fields(1)
                  TempStr = m_KeyAccountDetail.CUSTOMER_NAME & AddStringFrontEnd(m_KeyAccountDetail.BRANCH_NAME, "(", ")")
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------------------------------'
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmounts, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & INVOICE_DOCTYPE)
                  SumAmount = SumAmount + TempBd.TOTAL_AMOUNT
                  SumPrice = SumPrice + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                     
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmounts, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RECEIPT1_DOCTYPE)
                  SumAmount = SumAmount + TempBd.TOTAL_AMOUNT
                  SumPrice = SumPrice + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmounts, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RETURN_DOCTYPE)
                  SumReturnAmount = SumReturnAmount + TempBd.TOTAL_AMOUNT
                  SumReturnPrice = SumReturnPrice + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------------------------------'
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----------ของ TOP--------------------'
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountExs, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & INVOICE_DOCTYPE)
                  SumAmountTop = SumAmountTop + TempBd.TOTAL_AMOUNT
                  SumPriceTop = SumPriceTop + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                           
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountExs, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RECEIPT1_DOCTYPE)
                  SumAmountTop = SumAmountTop + TempBd.TOTAL_AMOUNT
                  SumPriceTop = SumPriceTop + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountExs, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RETURN_DOCTYPE)
                  SumReturnAmountTop = SumReturnAmountTop + TempBd.TOTAL_AMOUNT
                  SumReturnPriceTop = SumReturnPriceTop + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----------ของ TOP--------------------'
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------เรียบเทียบ------------------------'
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountCompares, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & INVOICE_DOCTYPE)
                  SumAmountCompare = SumAmountCompare + TempBd.TOTAL_AMOUNT
                  SumPriceCompare = SumPriceCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                     
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountCompares, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RECEIPT1_DOCTYPE)
                  SumAmountCompare = SumAmountCompare + TempBd.TOTAL_AMOUNT
                  SumPriceCompare = SumPriceCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountCompares, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RETURN_DOCTYPE)
                  SumReturnAmountComPare = SumReturnAmountComPare + TempBd.TOTAL_AMOUNT
                  SumReturnPriceCompare = SumReturnPriceCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------เรียบเทียบ------------------------'
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-------เรียบเทียบ----ของ TOP--------------------'
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountExCompares, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & INVOICE_DOCTYPE)
                  SumAmountTopCompare = SumAmountTopCompare + TempBd.TOTAL_AMOUNT
                  SumPriceTopCompare = SumPriceTopCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                           
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountExCompares, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RECEIPT1_DOCTYPE)
                  SumAmountTopCompare = SumAmountTopCompare + TempBd.TOTAL_AMOUNT
                  SumPriceTopCompare = SumPriceTopCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  
                  Set TempBd = GetObject("CBillingDoc", m_SaleAparAmountExCompares, Emp.EMP_CODE & "-" & m_KeyAccountDetail.CUSTOMER_CODE & "-" & m_KeyAccountDetail.BRANCH_CODE & "-" & RETURN_DOCTYPE)
                  SumReturnAmountTopComPare = SumReturnAmountTopComPare + TempBd.TOTAL_AMOUNT
                  SumReturnPriceTopCompare = SumReturnPriceTopCompare + TempBd.TOTAL_PRICE - TempBd.DISCOUNT_AMOUNT - TempBd.EXT_DISCOUNT_AMOUNT
                  '------------------------------------ จำนวน INVOICE ใบเสร็จขายสด รับคืน-----เรียบเทียบ------ของ TOP--------------------'
                  k = 1
                  If mcolParam("SHOW_AMOUNT") = 1 Then
                     k = k + 1
                     Amt = SumAmount + SumAmountTop
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_PRICE") = 1 Then
                     k = k + 1
                     Amt = SumPrice + SumPriceTop
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                     k = k + 1
                     Amt = SumReturnAmount + SumReturnAmountTop
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                     k = k + 1
                     Amt = SumReturnPrice + SumReturnPriceTop
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_AMOUNT") = 1 Then
                     k = k + 1
                     Amt = SumAmountCompare + SumAmountTopCompare
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_PRICE") = 1 Then
                     k = k + 1
                     Amt = SumPriceCompare + SumPriceTopCompare
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                     k = k + 1
                     Amt = SumReturnAmountComPare + SumReturnAmountTopComPare
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                     k = k + 1
                     Amt = SumReturnPriceCompare + SumReturnPriceTopCompare
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  '''''''''''''
                  If mcolParam("SHOW_AMOUNT") = 1 Then
                     k = k + 1
                     Amt = (SumAmount + SumAmountTop) - (SumAmountCompare + SumAmountTopCompare)
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_PRICE") = 1 Then
                     k = k + 1
                     Amt = (SumPrice + SumPriceTop) - (SumPriceCompare + SumPriceTopCompare)
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_RETURN_AMOUNT") = 1 Then
                     k = k + 1
                     Amt = (SumReturnAmount + SumReturnAmountTop) - (SumReturnAmountComPare + SumReturnAmountTopComPare)
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  
                  If mcolParam("SHOW_RETURN_PRICE") = 1 Then
                     k = k + 1
                     Amt = (SumReturnPrice + SumReturnPriceTop) - (SumReturnPriceCompare + SumReturnPriceTopCompare)
                     Set HeadCf = m_Headers1.Fields(k)
                     TempStr = FormatNumberToNull(Amt)
                     Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                     Call m_Details1.AddField(BodyCf)
                     Total1(k) = Total1(k) + Amt
                  End If
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  
               Next m_KeyAccountDetail
               If m_KeyAccount.KeyAccountDetail.Count > 0 Then
                  Call GenerateFooter("รวม" & "" & "", Total1, m_Details1, 0)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
                  
                  Set m_KeyAccount = Nothing
                  Set m_Rs = Nothing
                  
                  Vsp = "หมาย เหตุ     :"
                  Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
                  Vsp = ""
                  Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
                  Vsp = ""
                  Call Vsp.DrawLine(Vsp.CurrentX, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
               End If
            End If
         End If
      Next Emp
            
'      Call GenerateFooter("      สุทธิ" & "" & "", Total2, m_Details1, 0)
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
      
      Call GenerateEndDoc
      
      Set m_SaleAmounts = Nothing
      Set m_SaleAmountExs = Nothing
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
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
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = iSize
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   
   If CanPrintHeaderKeyAccount Then
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
Private Function GenerateCompareString(Text As String, Value, ValueCompare) As String
   If Round(Value, 2) <> Round(ValueCompare, 2) Then
      GenerateCompareString = Text & "" & FormatNumberToNull(MyDiff((Value - ValueCompare) * 100, Value))
   Else
      GenerateCompareString = Text
   End If
End Function
Private Function CheckInSaleCode(Code As String) As Boolean
   CheckInSaleCode = True
   If mcolParam("FROM_SALE_CODE") > Code And Len(mcolParam("FROM_SALE_CODE")) > 0 Then
      CheckInSaleCode = False
   End If
   If mcolParam("TO_SALE_CODE") < Code And Len(mcolParam("TO_SALE_CODE")) > 0 Then
      CheckInSaleCode = False
   End If
End Function
