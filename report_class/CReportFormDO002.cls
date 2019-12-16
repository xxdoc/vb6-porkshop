VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFormDO002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CReportInterface

Private Const MODULE_NAME = "CReportFormDO002"
Private MODULE_DESC As String

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
Private Const HEADER_SIZE = 16
Private Const DETAIL_SIZE = 16

Private Const TITLE_FORMAT = "<52|<3|<5;"
Private Const HEADER_FORMAT = "^1|^6|^2|^2|^3;"
Private Const DETAIL_FORMAT = "^1|<6|>2|>2|>3;"
Private Const FOOTER_FORMAT = "^1|>6|>2|>2|>3;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As CName
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig

Private m_cData As CBillingDoc

Private m_ImageWidth As Double
Private m_ImageHeight As Double
Private m_BillingDocIDColl As Collection
Private DriverIdCOll  As Collection
Private TranSportorIdCOll  As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_ReportConfig = New CReportConfig
   Set m_BillingDocIDColl = New Collection
   Set DriverIdCOll = New Collection
   Set TranSportorIdCOll = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_ReportConfig = Nothing
   Set m_BillingDocIDColl = Nothing
   Set DriverIdCOll = Nothing
   Set TranSportorIdCOll = Nothing
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

Private Sub PrintPOFooter(alngX() As Long, BD As CBillingDoc, Optional ShowSummary As Boolean = True)
      Call GenerateDocFooter(2, Vsp, mcolParam, tbNone, BD, ShowSummary)
End Sub
Private Sub PrintPOHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempBorder As TableBorderSettings
Dim TempMarginLeft As Double

   TempMarginLeft = Vsp.MarginLeft
   Vsp.MarginLeft = 567 * 2 - m_ReportConfig.GetFieldValue("DUMMY_OFFSET")
   
   Vsp.FontSize = 14.25
   
   Call GenerateDocHeader(1, Vsp, mcolParam, tbNone, m_ReportConfig.GetFieldValue("HEAD_OFFSET"))
    
    
End Sub

Private Sub printHeader()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdt As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Call PrintPOHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   
      
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
   
End Sub
Private Function initDoc() As Boolean
Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   
   Call LoadDisTinctBillingDocID(m_BillingDocIDColl, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), PatchWildCard(mcolParam("DOCUMENT_NO_SEARCH")), , , , mcolParam("ORDER_BY"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")", mcolParam("DRIVER_ID"), mcolParam("TRANSPORTOR_ID"))
   
   Vsp.PaperSize = pprLetter
    m_ImageWidth = 19.5 * 567
    m_ImageHeight = 14.36 * 567
   
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = "1cm"
   Vsp.MarginFooter = "1cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginTop = "0cm"
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   MODULE_DESC = MapText("ใบส่งสินค้า")
   
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)
   
   mdblWid = Vsp.PageWidth - (2 * m_ReportConfig.GetFieldValue("DUMMY_OFFSET"))
   
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
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Poi As CDocItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim j As Long
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
Dim AskMessage As VbMsgBoxResult
Dim TempString As String
Dim TempBillingDocID  As CBillingDoc
   RName = "genDoc"
   
   TempString = "ต้องการพิมพ์ทันทีหรือไม่? (YES พิมพ์ทันที   NO แสดงผลรายงานก่อน  CANCEL ออก) "
   AskMessage = MsgBox(TempString, vbYesNoCancel, glbErrorLog.MsgBoxTitle)
   If AskMessage = vbCancel Then
      frmReport.AutoPrintMode = True
      Exit Function
   End If
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   If Not initDoc Then Exit Function
   
   If m_BillingDocIDColl.Count <= 0 Then
         mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
         Exit Function
   End If
   
   If AskMessage = vbNo Then
      Vsp.Preview = True
   Else
      Vsp.Preview = False
   End If
   
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   Call StartExportFile(Vsp)
      
   mblnHeader = False
   Vsp.StartDoc
   I = 0
   
   For Each TempBillingDocID In m_BillingDocIDColl
      I = I + 1
      Set Rs = New ADODB.Recordset
      Set m_cData = New CBillingDoc
      m_cData.BILLING_DOC_ID = TempBillingDocID.BILLING_DOC_ID
      m_cData.QueryFlag = 1
      If Not glbDaily.QueryBillingDocOnlyDo(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
         mstrErrMsg = "Query Error."
         Exit Function
      End If
      
      Call m_cData.PopulateFromRS(1, Rs)
      Call mcolParam.add(m_cData.DOCUMENT_NO, "DOCUMENT_NO")
      Call mcolParam.add(m_cData.DOCUMENT_DATE, "DOCUMENT_DATE")
      Call mcolParam.add(m_cData.APAR_CODE, "CUSTOMER_CODE")
      Call mcolParam.add(m_cData.APAR_NAME, "CUSTOMER_NAME")
      Call mcolParam.add(m_cData.CREDIT, "CREDIT")
      Call mcolParam.add(m_cData.DUE_DATE, "DUE_DATE")
      Call mcolParam.add(m_cData.TAX_ID, "TAX_ID")
      
      Call mcolParam.add(m_cData.SALE_CODE, "SALE_CODE")
      Call mcolParam.add(m_cData.SALE_NAME, "SALE_NAME")
      Call mcolParam.add(m_cData.CUS_PO, "CUS_PO")
      Call mcolParam.add(m_cData.CUSTOMER_BRANCH_NAME, "CUSTOMER_BRANCH_NAME")
      
      If m_cData.BILLING_ADDRESS_ID > 0 Then
         Set Addr = New CAddress
         Call Addr.SetFieldValue("ADDRESS_ID", m_cData.BILLING_ADDRESS_ID)
         Call Addr.QueryData(1, Rs, iCount)
         If Not Rs.EOF Then
            Call Addr.PopulateFromRS(1, Rs)
            Call mcolParam.add(Addr.PackAddress, "CUSTOMER_ADDRESS")
            Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "PHONE")
            Call mcolParam.add(Addr.GetFieldValue("FAX1"), "FAX")
         End If
         Set Addr = Nothing
      Else
         Call mcolParam.add("", "CUSTOMER_ADDRESS")
         Call mcolParam.add("", "PHONE")
         Call mcolParam.add("", "FAX")
      End If
   
      If m_cData.BRANCH_ADDRESS > 0 Then
         Set Addr = New CAddress
         Call Addr.SetFieldValue("ADDRESS_ID", m_cData.BRANCH_ADDRESS)
         Call Addr.QueryData(1, Rs, iCount)
         If Not Rs.EOF Then
            Call Addr.PopulateFromRS(1, Rs)
             Call mcolParam.add(Addr.PackAddress, "BRANCH_ADDRESS")
            Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "BRANCH_ADDRESS_PHONE")
            Call mcolParam.add(Addr.GetFieldValue("FAX1"), "BRANCH_ADDRESS_FAX")
            Call mcolParam.add(Addr.GetFieldValue("SHOW_LOCATION_FLAG"), "SHOW_LOCATION_FLAG")
         End If
         Set Addr = Nothing
      Else
         Call mcolParam.add("", "BRANCH_ADDRESS")
         Call mcolParam.add("", "BRANCH_ADDRESS_PHONE")
         Call mcolParam.add("", "BRANCH_ADDRESS_FAX")
         Call mcolParam.add("", "SHOW_LOCATION_FLAG")
      End If
      
      If I > 1 Then
         Vsp.NewPage
      Else
         Call printHeader
         mblnHeader = True
      End If
      
      
      
      TempBorder = tbNone
      
      Call GenerateDocBody(1, Vsp, mcolParam, m_cData, tbNone)
      
      Call PrintPOFooter(alngX, m_cData)
            
      '--------------      Clear mColparam ----------------'
      Call mcolParam.Remove("DOCUMENT_NO")
      Call mcolParam.Remove("DOCUMENT_DATE")
      Call mcolParam.Remove("CUSTOMER_CODE")
      Call mcolParam.Remove("CUSTOMER_NAME")
      Call mcolParam.Remove("CUSTOMER_BILL_NAME")
      Call mcolParam.Remove("CREDIT")
      Call mcolParam.Remove("DUE_DATE")
      
      Call mcolParam.Remove("SALE_CODE")
      Call mcolParam.Remove("SALE_NAME")
      Call mcolParam.Remove("CUS_PO")
      Call mcolParam.Remove("TAX_ID")
      Call mcolParam.Remove("CUSTOMER_BRANCH_NAME")
      
      Call mcolParam.Remove("CUSTOMER_ADDRESS")
      Call mcolParam.Remove("PHONE")
      Call mcolParam.Remove("FAX")
      
      Call mcolParam.Remove("BRANCH_ADDRESS")
      Call mcolParam.Remove("BRANCH_ADDRESS_PHONE")
      Call mcolParam.Remove("BRANCH_ADDRESS_FAX")
      Call mcolParam.Remove("SHOW_LOCATION_FLAG")
      '--------------      Clear mColparam ----------------'
      
      Set Rs = Nothing
      Set m_cData = Nothing
   Next TempBillingDocID
   
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   
   If AskMessage = vbYes Then
      Vsp.PrintDoc (True)
      frmReport.AutoPrintMode = True
   End If
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.Description
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   Set Rs = Nothing
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
   
'   Vsp.DrawPicture mcolParam("PICTURE"), "0.75in", "0.75in", "100%", "100%"
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage & " / " & glbUser.USER_NAME)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
      
   
   If mblnHeader Then
      Call printHeader
   End If
   
   'mdY = Vsp.CurrentY
   
   'Vsp.TextAlign = talnT
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

Public Sub GenerateDocFooter(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, BD As CBillingDoc, Optional ShowSummary As Boolean = True)
Dim TempX As Long
Dim TempY As Long
Dim Amt As Double
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim TempFont As Byte
Dim TempString As String
Dim Mr1 As CMasterRef
Dim Mr2 As CMasterRef
   Vsp.FontBold = True
   Vsp.TableBorder = TempBorder
   
   TempX = Vsp.CurrentX
   TempY = Vsp.CurrentY
   TempFont = Vsp.FontSize
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "2.56cm"
    Vsp.TableCell(tcColWidth, , 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, , 2) = "8cm"
    Vsp.TableCell(tcColWidth, , 3) = "4.5cm"
    Vsp.TableCell(tcColWidth, , 4) = "1.3cm"
    Vsp.TableCell(tcColWidth, , 5) = "3cm"
   
   Vsp.FontSize = 12
   If mcolParam("SHOW_LOCATION_FLAG") = "Y" Then
      If ShowSummary Then
         Vsp.TableCell(tcText, 1, 1) = "สถานที่ส่งของ"
         Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
         
         Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME") & " (สาขา " & BD.CUSTOMER_BRANCH_NAME & ")" & vbCrLf & mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX")
         Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
      End If
   End If
   Vsp.TableBorder = TempBorder
   Vsp.EndTable
   
   
   Vsp.CurrentX = TempX
   Vsp.CurrentY = TempY
   Vsp.FontSize = TempFont
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 5
   Vsp.TableCell(tcRows) = 5
    Vsp.TableCell(tcRowHeight) = "0.64cm"
    Vsp.TableCell(tcColWidth, , 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, , 2) = "8cm"
    Vsp.TableCell(tcColWidth, , 3) = "4.5cm"
    Vsp.TableCell(tcColWidth, , 4) = "1.3cm"
    Vsp.TableCell(tcColWidth, , 5) = "3cm"
   
   If ShowSummary Then
      Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
      Vsp.TableCell(tcText, 1, 5) = FormatNumber(BD.TOTAL_PRICE)
   
      Sum1 = Val(FormatNumber(BD.DISCOUNT_AMOUNT, , False)) + Val(FormatNumber(BD.EXT_DISCOUNT_AMOUNT, , False))
   
      If BD.EXT_DISCOUNT_PERCENT > 0 Then
         Vsp.TableCell(tcAlign, 2, 4) = taCenterMiddle
         Vsp.TableCell(tcText, 2, 4) = BD.EXT_DISCOUNT_PERCENT & " %"
      End If
   
      Vsp.TableCell(tcAlign, 2, 5) = taRightMiddle
      Vsp.TableCell(tcText, 2, 5) = FormatNumber(Sum1)
      
      Sum2 = Val(FormatNumber(BD.TOTAL_PRICE, , False)) - Sum1
   
      Vsp.TableCell(tcAlign, 3, 5) = taRightMiddle
      Vsp.TableCell(tcText, 3, 5) = FormatNumber(Sum2)
   
      If BD.VAT_PERCENT > 0 Then
         Vsp.TableCell(tcAlign, 4, 4) = taCenterMiddle
         Vsp.TableCell(tcText, 4, 4) = BD.VAT_PERCENT & " %"
      End If
   
      Vsp.TableCell(tcAlign, 4, 5) = taRightMiddle
      Vsp.TableCell(tcText, 4, 5) = FormatNumber(BD.VAT_AMOUNT)
      
      Amt = Sum2 + Val(FormatNumber(BD.VAT_AMOUNT, , False))
      
      Vsp.TableCell(tcAlign, 5, 2) = taRightMiddle
      Vsp.TableCell(tcText, 5, 2) = "(" & ThaiBaht(Amt) & ")"
      
      Vsp.TableCell(tcAlign, 5, 5) = taRightMiddle
      Vsp.TableCell(tcText, 5, 5) = FormatNumber(Amt)
   End If
   Vsp.TableBorder = TempBorder
   Vsp.EndTable
      
   '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 2.3 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
    
      '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 3
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
   Vsp.TableCell(tcColWidth, 1, 1) = "9.3cm"
   Vsp.TableCell(tcColWidth, 1, 2) = "3.9cm"
   Vsp.TableCell(tcColWidth, 1, 3) = "6.2cm"
    
'   Set Mr1 = GetObject("CMasterRef", DriverIdCOll, Trim(Str(BD.DRIVER_ID)))
'   Set Mr2 = GetObject("CMasterRef", TranSportorIdCOll, Trim(Str(BD.TRANSPORTOR_ID)))
   
   Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   Vsp.TableCell(tcText, 1, 2) = Mr1.KEY_NAME & "  " & Mr2.KEY_NAME
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
    
   '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
    Vsp.TableBorder = TempBorder
   
    Vsp.EndTable
'   '============

End Sub
Public Sub GenerateDocHeader(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, HeaderOffset As Double)
Dim TempX As Long
Dim TempY As Long
   
   Vsp.FontBold = True
   
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = 4.1 * 567 + HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 2
    Vsp.TableCell(tcRowHeight) = 1 * 567
    Vsp.TableCell(tcColWidth, , 1) = "15.5cm"
    Vsp.TableCell(tcColWidth, , 2) = "3.8cm"
   
    Vsp.TableCell(tcText, 1, 2) = mcolParam("DOCUMENT_NO")
    Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   
   Vsp.TableCell(tcText, 2, 2) = DateToStringExtEx2(mcolParam("DOCUMENT_DATE"))
    Vsp.TableCell(tcAlign, 2, 2) = taLeftMiddle
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============

   '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
    Vsp.TableBorder = TempBorder

    Vsp.EndTable
'   '============
   
   TempX = Vsp.CurrentX
   TempY = Vsp.CurrentY
'   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 2
    Vsp.TableCell(tcRowHeight) = 1.2 * 567
    Vsp.TableCell(tcColWidth, , 1) = "3cm"
    Vsp.TableCell(tcColWidth, , 2) = "9.2cm"
    Vsp.TableCell(tcColWidth, , 3) = "3.3cm"
    Vsp.TableCell(tcColWidth, , 4) = "3.8cm"
      
   Vsp.TableCell(tcText, 1, 4) = mcolParam("CUSTOMER_CODE")
   Vsp.TableCell(tcAlign, 1, 4) = taLeftMiddle
   Vsp.TableCell(tcText, 2, 4) = mcolParam("CUS_PO")
    Vsp.TableCell(tcAlign, 2, 4) = taLeftMiddle
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
'   '============

   Vsp.CurrentX = TempX
   Vsp.CurrentY = TempY
'   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 2.4 * 567
    Vsp.TableCell(tcColWidth, , 1) = "3cm"
    Vsp.TableCell(tcColWidth, , 2) = "9.2cm"
    Vsp.TableCell(tcColWidth, , 3) = "3.3cm"
    Vsp.TableCell(tcColWidth, , 4) = "3.8cm"
    
    If Len(mcolParam("CUSTOMER_BILL_NAME")) > 0 Then
      Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_BILL_NAME") & vbCrLf & mcolParam("CUSTOMER_ADDRESS") & vbCrLf & "TAX ID       " & mcolParam("TAX_ID")
   Else
      Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME") & vbCrLf & mcolParam("CUSTOMER_ADDRESS") & vbCrLf & "TAX ID       " & mcolParam("TAX_ID")
   End If
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
'   '============

   
   '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
    Vsp.TableBorder = TempBorder

    Vsp.EndTable
'   '============


'   '============

   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.6 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "6.1cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "4.7cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "4.7cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "3.8cm"
    Vsp.TableBorder = TempBorder
   
    Vsp.TableCell(tcText, 1, 1) = mcolParam("SALE_CODE") & "-" & mcolParam("SALE_NAME")
    Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
   
    Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_BRANCH_NAME")
    Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
   
      Vsp.TableCell(tcText, 1, 3) = mcolParam("CREDIT")
      Vsp.TableCell(tcAlign, 1, 3) = taCenterMiddle
   
      Vsp.TableCell(tcText, 1, 4) = DateToStringExtEx2(mcolParam("DUE_DATE"))
      Vsp.TableCell(tcAlign, 1, 4) = taCenterMiddle
      
      
    Vsp.EndTable
'   '============

'   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.2 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
    Vsp.TableBorder = TempBorder

    Vsp.EndTable
'   '============
End Sub

Public Sub GenerateDocBody(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, BD As CBillingDoc, TempBorder As TableBorderSettings)
Dim Poi As CDocItem
Dim I As Long
Dim j As Long
Dim alngX() As Long
   
   I = 0
   j = 0
   
   ''debug.print (Vsp.FontSize)
   
   For Each Poi In BD.DocItems
      I = I + 1
      j = j + 1

      If j > 9 Then
         j = 1
         Call PrintPOFooter(alngX, m_cData, False)
         Vsp.NewPage
      End If

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 5
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight) = "0.7cm"
      Vsp.TableBorder = TempBorder

       Vsp.TableCell(tcColWidth, , 1) = "1.3cm"
       Vsp.TableCell(tcColWidth, , 2) = "8.7cm"
       Vsp.TableCell(tcColWidth, , 3) = "3.4cm"
       Vsp.TableCell(tcColWidth, , 4) = "2.9cm"
       Vsp.TableCell(tcColWidth, , 5) = "3.0cm"


       Vsp.TableCell(tcAlign, 1, 1) = taCenterMiddle
       Vsp.TableCell(tcText, 1, 1) = I
       
       Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 2) = Poi.ShowDescCode & " " & Poi.ShowDescText

       Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
       Vsp.TableCell(tcText, 1, 3) = FormatNumber(MyDiff(Poi.GetFieldValue("ITEM_AMOUNT"), Poi.GetFieldValue("UNIT_MULTIPLE"))) & " " & Poi.GetFieldValue("UNIT_TRAN_NAME")

       Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
       Vsp.TableCell(tcText, 1, 4) = FormatNumber(Poi.GetFieldValue("AVG_PRICE") * Poi.GetFieldValue("UNIT_MULTIPLE"))

       Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
       Vsp.TableCell(tcText, 1, 5) = FormatNumber(Poi.GetFieldValue("TOTAL_PRICE"))

       Vsp.EndTable
   Next Poi

   For I = 1 To (9 - j)
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight) = "0.7cm"
       Vsp.TableCell(tcColWidth, , 1) = "19.3cm"
      Vsp.EndTable
   Next I

      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight) = "0.1cm"
       Vsp.TableCell(tcColWidth, , 1) = "19.3cm"
      Vsp.EndTable
      
End Sub
