VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportFormRv0001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CReportInterface

Private Const MODULE_NAME = "CReportFormRv001"
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
Private DriverIdCOll As Collection
Private TranSportorIdCOll As Collection
Private TempDocumentRef As String
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_ReportConfig = New CReportConfig
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
   If mcolParam("PREVIEW_TYPE") = 2 Then
      CReportInterface_Preview = genDoc(False)
   Else
      CReportInterface_Preview = genDoc(True)
   End If
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

Private Sub PrintPOFooter(alngX() As Long, BD As CBillingDoc)
      Call GenerateDocFooter(2, Vsp, mcolParam, tbNone, BD)
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
   Vsp.MarginLeft = (567 * 2) - m_ReportConfig.GetFieldValue("DUMMY_OFFSET")
   
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
'   strFormat = VSP_CalTable(HEADER_FORMAT, mdblWid, alngX)
'   strPrint = "ลำดับ|สินค้าและบริการ|จำนวน|ราคาต่อหน่วย|จำนวนเงิน;"
'   Call Vsp.AddTable(strFormat, "", strPrint)
   
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
      
   Vsp.PaperSize = pprA4
   
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
Private Sub GenerateDoItem(Src As Collection, Dest As Collection)
'Dim Ri As CReceiptItem
'Dim Di As CDoItem
'
'   For Each Ri In Src
'      Set Di = New CDoItem
'      Di.ITEM_AMOUNT = 1
'      Di.TOTAL_PRICE = Ri.RECEIPT_ITEM_AMOUNT
'      Di.AVG_PRICE = Ri.RECEIPT_ITEM_AMOUNT
'      Di.DISPLAY_ID = 1
'      Di.DISCOUNT_AMOUNT = Ri.DISCOUNT_AMOUNT
'      If m_cData.RECEIPT_TYPE = 3 Then
'         Di.ITEM_DESC = "อ้างถึงเอกสาร " & Ri.DOCUMENT_NO
'      ElseIf m_cData.RECEIPT_TYPE = 5 Then
'         Di.ITEM_DESC = "อ้างถึงเอกสาร " & Ri.DOCUMENT_NO
'      End If
'      Call Dest.Add(Di)
'      Set Di = Nothing
'   Next Ri
End Sub
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

   RName = "genDoc"
Dim Rcp As CRcpCnDn_Item
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
      
   Call m_cData.PopulateFromRS(1, Rs)
   If m_cData.DOCUMENT_TYPE = RECEIPT2_DOCTYPE Or m_cData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
      Call mcolParam.add(m_cData.DOCUMENT_NO, "DOCUMENT_NO")
      Call mcolParam.add(m_cData.APAR_NAME, "CUSTOMER_NAME")
   ElseIf m_cData.DOCUMENT_TYPE = RECEIPT3_DOCTYPE Then
      Call mcolParam.add(m_cData.NOTE, "CUSTOMER_NAME")
   End If
   Call mcolParam.add(m_cData.DOCUMENT_DATE, "DOCUMENT_DATE")
   'Call mcolParam.add(m_cData.RCP_CASH_TRAN - m_cData.FEE_AMOUNT, "RV_AMOUNT")
   Call mcolParam.add(m_cData.RCP_CASH_TRAN, "RV_AMOUNT")      ' หักค่าธรรมเนียม
   
'   For Each Rcp In m_cData.RcpCnDnItems
'      TempDocumentRef = TempDocumentRef & " " & Rcp.GetFieldValue("DOC_NO")
'   Next Rcp
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   Call StartExportFile(Vsp)
      
   Vsp.StartDoc
   
   TempBorder = tbNone
   
   Call GenerateDocBody(1, Vsp, mcolParam, m_cData, tbNone)
      
   Call PrintPOFooter(alngX, m_cData)
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   
   If mcolParam("PREVIEW_TYPE") = 2 Then
      Vsp.PrintDoc (True)
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
      
   Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.4cm", "20.4cm", "27.8cm", 10
   'Vsp.DrawPicture mcolParam("BACK_GROUND"), Vsp.MarginLeft, Vsp.MarginTop, m_ImageWidth, m_ImageHeight, 10
   
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

Public Sub GenerateDocFooter(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, BD As CBillingDoc)
Dim TempX As Long
Dim TempY As Long
Dim Amt As Double
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim TempFont As Byte
Dim TempString As String
Dim Mr  As CMasterRef
Dim SumDr As Double
Dim SumCr As Double
   
   Vsp.FontBold = True
   Vsp.TableBorder = TempBorder
   
   TempX = Vsp.CurrentX
   TempY = Vsp.CurrentY
   TempFont = Vsp.FontSize
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 7
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = "1.9cm"
    Vsp.TableCell(tcColWidth, , 1) = "0.8cm"
    Vsp.TableCell(tcColWidth, , 2) = "9cm"
    Vsp.TableCell(tcColWidth, , 3) = "2.1cm"
    Vsp.TableCell(tcColWidth, , 4) = "2.8cm"
    Vsp.TableCell(tcColWidth, , 5) = "0.75cm"
   Vsp.TableCell(tcColWidth, , 6) = "2.8cm"
   Vsp.TableCell(tcColWidth, , 7) = "0.75cm"
   
   SumDr = BD.RCP_CASH_TRAN + BD.FEE_AMOUNT + BD.SUBTRACT_AMOUNT  '
   SumCr = BD.PAID_AMOUNT - BD.CREDIT_AMOUNT + BD.DEBIT_AMOUNT + BD.ADDITION_AMOUNT
   
   If SumDr >= 0 Then
      Vsp.TableCell(tcText, 1, 2) = "(" & ThaiBaht(SumDr) & ")"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   Else
      Vsp.TableCell(tcText, 1, 2) = "(ลบ" & ThaiBaht(-SumDr) & ")"
      Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   End If
   
   Vsp.TableCell(tcText, 1, 4) = FormatNumber(SumDr)
   Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
   
   Vsp.TableCell(tcText, 1, 6) = FormatNumber(SumCr)
   Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
   
   Vsp.TableBorder = TempBorder
   Vsp.EndTable
   
   
End Sub
Private Sub GenerateDocHeader(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, HeaderOffset As Double)
Dim iSize As Integer
   Vsp.FontBold = True
   
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = (2.4 * 567) - HeaderOffset
    Vsp.TableCell(tcColWidth, 1, 1) = "19cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 1.5 * 567
    Vsp.TableCell(tcColWidth, , 1) = "17cm"
    Vsp.TableCell(tcColWidth, , 2) = "2cm"
   
    Vsp.TableCell(tcText, 1, 2) = mcolParam("DOCUMENT_NO")
    Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============

   '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "14cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "5cm"
    
    Vsp.TableCell(tcText, 1, 2) = DateToStringExtEx2(mcolParam("DOCUMENT_DATE"))
    Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
    
    Vsp.TableBorder = TempBorder

    Vsp.EndTable
'   '============

'   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight) = 0.9 * 567
    Vsp.TableCell(tcColWidth, , 1) = "4cm"
    Vsp.TableCell(tcColWidth, , 2) = "15cm"
    
   Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME")
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
'   '============
   
   '   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "4cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "15cm"
    
   Vsp.TableCell(tcText, 1, 2) = "สินค้า"
   Vsp.TableCell(tcAlign, 1, 2) = taLeftMiddle
   
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
'   '============
   
'   '============
   iSize = Vsp.FontSize
   Vsp.FontSize = 10
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 2
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 2.4 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "3.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "15.5cm"
      
   Vsp.TableCell(tcText, 1, 2) = "" 'TempDocumentRef         'เอกสารประกอบ
   Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
   
   Vsp.TableBorder = TempBorder
   Vsp.EndTable
'   '============
   
   Vsp.FontSize = iSize
   
'   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 4
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 0.8 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "2.5cm"
    Vsp.TableCell(tcColWidth, 1, 2) = "3.5cm"
    Vsp.TableCell(tcColWidth, 1, 3) = "2cm"
    Vsp.TableCell(tcColWidth, 1, 4) = "11cm"
    
    Vsp.TableCell(tcText, 1, 2) = FormatNumber(mcolParam("RV_AMOUNT"))                        'เอกสารประกอบ
    Vsp.TableCell(tcAlign, 1, 2) = taCenterTop
    
    If mcolParam("RV_AMOUNT") >= 0 Then
      Vsp.TableCell(tcText, 1, 4) = "(" & ThaiBaht(mcolParam("RV_AMOUNT")) & ")"                        'เอกสารประกอบ
      Vsp.TableCell(tcAlign, 1, 4) = taCenterTop
    Else
      Vsp.TableCell(tcText, 1, 4) = "(ลบ" & ThaiBaht(-mcolParam("RV_AMOUNT")) & ")"                        'เอกสารประกอบ
      Vsp.TableCell(tcAlign, 1, 4) = taCenterTop
    End If
    
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
'   '============

   '-------------------------------------------------------------------------------- ข้อมูล รับเป็น อะไร เงินสด เช็ค โอน
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 3.9 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19cm"
   Vsp.TableBorder = TempBorder
   Vsp.EndTable
   '-------------------------------------------------------------------------------- ข้อมูล รับเป็น อะไร เงินสด เช็ค โอน
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
    Vsp.TableCell(tcRowHeight, 1) = 1.6 * 567
    Vsp.TableCell(tcColWidth, 1, 1) = "19cm"
   Vsp.TableBorder = TempBorder
   Vsp.EndTable
End Sub
Public Sub GenerateDocBody(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, BD As CBillingDoc, TempBorder As TableBorderSettings)
Dim Poi As CCashTran
Dim Bs As CBillingSubTract
Dim Ba As CBillingAddition
Dim I As Long
Dim j As Long
Dim BankAccountColl As Collection
Dim AddColl As Collection
Dim SubColl As Collection
Dim Mr As CMasterRef
   
   Set BankAccountColl = New Collection
   Set AddColl = New Collection
   Set SubColl = New Collection
   
   Call LoadMaster(Nothing, BankAccountColl, , , MASTER_BANK_ACCOUNT)
   Call LoadMaster(Nothing, AddColl, , , MASTER_ADDITION)
   Call LoadMaster(Nothing, SubColl, , , MASTER_SUBTRACT)
   I = 0
   j = 0
   For Each Poi In BD.Payments
      I = I + 1
      j = j + 1

      If j > 12 Then
         j = 1
         Vsp.NewPage
      End If
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.7cm"
      Vsp.TableBorder = TempBorder

       Vsp.TableCell(tcColWidth, , 1) = "9.8cm"
       Vsp.TableCell(tcColWidth, , 2) = "2.1cm"
       Vsp.TableCell(tcColWidth, , 3) = "2.8cm"
       Vsp.TableCell(tcColWidth, , 4) = "0.75cm"
       Vsp.TableCell(tcColWidth, , 5) = "2.8cm"
       Vsp.TableCell(tcColWidth, , 6) = "0.75cm"
      
       Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
        Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
       If Poi.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
         Vsp.TableCell(tcText, 1, 1) = "เงินสด"
         Vsp.TableCell(tcText, 1, 2) = ""
       ElseIf Poi.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Or Poi.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT Then
         Vsp.TableCell(tcText, 1, 1) = "เช็ค"
         Vsp.TableCell(tcText, 1, 2) = ""
       ElseIf Poi.GetFieldValue("PAYMENT_TYPE") = BANKTRF_PMT Then
         Set Mr = GetObject("CMasterRef", BankAccountColl, Trim(Str(Poi.GetFieldValue("BANK_ACCOUNT"))))
          Vsp.TableCell(tcText, 1, 1) = "ธ. " & Poi.GetFieldValue("BANK_NAME") & "-" & Poi.GetFieldValue("BRANCH_NAME") & " ( " & Poi.GetFieldValue("ACCOUNT_NAME") & " )"
         Vsp.TableCell(tcText, 1, 2) = Mr.SHORT_CODE
       End If
       
       Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
       Vsp.TableCell(tcText, 1, 3) = FormatNumber(Poi.GetFieldValue("AMOUNT") - Poi.GetFieldValue("FEE_AMOUNT"))
       
       Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
       Vsp.TableCell(tcText, 1, 4) = ""

       Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
       Vsp.TableCell(tcText, 1, 5) = ""
      
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       Vsp.TableCell(tcText, 1, 6) = ""
       
       Vsp.EndTable
       
       '------------------ ค่าธรรมเนียม
       If Poi.GetFieldValue("FEE_AMOUNT") > 0 Then
         I = I + 1
         j = j + 1
   
         If j > 12 Then
            j = 1
            Vsp.NewPage
         End If
         
         Vsp.StartTable
         Vsp.TableCell(tcCols) = 6
         Vsp.TableCell(tcRows) = 1
         Vsp.TableCell(tcRowHeight) = "0.7cm"
         Vsp.TableBorder = TempBorder
   
          Vsp.TableCell(tcColWidth, , 1) = "9.8cm"
          Vsp.TableCell(tcColWidth, , 2) = "2.1cm"
          Vsp.TableCell(tcColWidth, , 3) = "2.8cm"
          Vsp.TableCell(tcColWidth, , 4) = "0.75cm"
          Vsp.TableCell(tcColWidth, , 5) = "2.8cm"
          Vsp.TableCell(tcColWidth, , 6) = "0.75cm"
         
          Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
          
          Vsp.TableCell(tcText, 1, 1) = "ค่าธรรมเนียม"
         Vsp.TableCell(tcText, 1, 2) = "720-4012"
          
          Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
          Vsp.TableCell(tcText, 1, 3) = FormatNumber(Poi.GetFieldValue("FEE_AMOUNT"))
          
          Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
          Vsp.TableCell(tcText, 1, 4) = ""
   
          Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
          Vsp.TableCell(tcText, 1, 5) = ""
         
          Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
          Vsp.TableCell(tcText, 1, 6) = ""
          
          Vsp.EndTable
       End If
   Next Poi
   
   For Each Bs In BD.BillingSubTracts
      I = I + 1
      j = j + 1

      If j > 12 Then
         j = 1
         Vsp.NewPage
      End If
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.7cm"
      Vsp.TableBorder = TempBorder
      
       Vsp.TableCell(tcColWidth, , 1) = "9.8cm"
       Vsp.TableCell(tcColWidth, , 2) = "2.1cm"
       Vsp.TableCell(tcColWidth, , 3) = "2.8cm"
       Vsp.TableCell(tcColWidth, , 4) = "0.75cm"
       Vsp.TableCell(tcColWidth, , 5) = "2.8cm"
       Vsp.TableCell(tcColWidth, , 6) = "0.75cm"
      
       Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 1) = Bs.GetFieldValue("SUBTRACT_NAME")
       
       Set Mr = GetObject("CMasterRef", SubColl, Trim(Str(Bs.GetFieldValue("SUBTRACT_ID"))))
       
       Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
       Vsp.TableCell(tcText, 1, 2) = Mr.SHORT_CODE
       
       Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
       Vsp.TableCell(tcText, 1, 3) = FormatNumber(Bs.GetFieldValue("ITEM_AMOUNT"))
      
       Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
       Vsp.TableCell(tcText, 1, 4) = ""

       Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
       Vsp.TableCell(tcText, 1, 5) = ""
      
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       Vsp.TableCell(tcText, 1, 6) = ""
       
       Vsp.EndTable
   Next Bs
   
   I = I + 1
   j = j + 1
   
   If j > 12 Then
      j = 1
      Vsp.NewPage
   End If
   
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 6
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight) = "0.7cm"
   Vsp.TableBorder = TempBorder
   
    Vsp.TableCell(tcColWidth, , 1) = "9.8cm"
    Vsp.TableCell(tcColWidth, , 2) = "2.1cm"
    Vsp.TableCell(tcColWidth, , 3) = "2.8cm"
    Vsp.TableCell(tcColWidth, , 4) = "0.75cm"
    Vsp.TableCell(tcColWidth, , 5) = "2.8cm"
    Vsp.TableCell(tcColWidth, , 6) = "0.75cm"
   
    Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
    Vsp.TableCell(tcText, 1, 1) = "                            ลูกหนี้การค้า"
    
    Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
    Vsp.TableCell(tcText, 1, 2) = "113-1100"
    
    Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
    Vsp.TableCell(tcText, 1, 3) = ""
   
    Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
    Vsp.TableCell(tcText, 1, 4) = ""

    Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
    Vsp.TableCell(tcText, 1, 5) = FormatNumber(BD.PAID_AMOUNT - BD.CREDIT_AMOUNT + BD.DEBIT_AMOUNT)
   
    Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
    Vsp.TableCell(tcText, 1, 6) = ""
    
    Vsp.EndTable
   
   
   For Each Ba In BD.BillingAdditions
      I = I + 1
      j = j + 1

      If j > 12 Then
         j = 1
         Vsp.NewPage
      End If
      
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 6
      Vsp.TableCell(tcRows) = 1
      Vsp.TableCell(tcRowHeight) = "0.7cm"
      Vsp.TableBorder = TempBorder
      
       Vsp.TableCell(tcColWidth, , 1) = "9.8cm"
       Vsp.TableCell(tcColWidth, , 2) = "2.1cm"
       Vsp.TableCell(tcColWidth, , 3) = "2.8cm"
       Vsp.TableCell(tcColWidth, , 4) = "0.75cm"
       Vsp.TableCell(tcColWidth, , 5) = "2.8cm"
       Vsp.TableCell(tcColWidth, , 6) = "0.75cm"
      
       Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
       Vsp.TableCell(tcText, 1, 1) = "                            " & Ba.GetFieldValue("ADDITION_NAME")
       
       Set Mr = GetObject("CMasterRef", AddColl, Trim(Str(Ba.GetFieldValue("ADDITION_ID"))))
       Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
       Vsp.TableCell(tcText, 1, 2) = Mr.SHORT_CODE
       
       Vsp.TableCell(tcAlign, 1, 3) = taRightMiddle
       Vsp.TableCell(tcText, 1, 3) = ""
      
       Vsp.TableCell(tcAlign, 1, 4) = taRightMiddle
       Vsp.TableCell(tcText, 1, 4) = ""

       Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
       Vsp.TableCell(tcText, 1, 5) = FormatNumber(Ba.GetFieldValue("ITEM_AMOUNT"))
      
       Vsp.TableCell(tcAlign, 1, 6) = taRightMiddle
       Vsp.TableCell(tcText, 1, 6) = ""
       
       Vsp.EndTable
   Next Ba
   
   For I = 1 To (12 - j)
      Vsp.StartTable
      Vsp.TableCell(tcCols) = 1
      Vsp.TableCell(tcRows) = 1
       Vsp.TableCell(tcRowHeight) = "0.7cm"
       Vsp.TableCell(tcColWidth, , 1) = "19cm"
      Vsp.EndTable
   Next I
   
   Set BankAccountColl = Nothing
   Set AddColl = Nothing
   Set SubColl = Nothing
End Sub
