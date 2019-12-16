VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalExportPromotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalExportPromotion"
Private MODULE_DESC  As String

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

Private Const HEADER_RECEIPT1 = "^1|^10|^3;"
Private Const HEADER_RECEIPT2 = "^1|^3|^3|^3|^3|^3|^3;"
Private Const HEADER_BILLS = "^1|^3|^3|^3|^3|^3|^3;"
Private Const HEADER_CNDN = "^1|^4|^4|^4;"

Private Const DETAIL_RECEIPT1 = "^1|<10|>3;"
Private Const DETAIL_RECEIPT2 = "^1|^3|^3|^3|>3|>3|>3;"
Private Const DETAIL_BILLS = "^1|^3|^3|^3|>3|>3|>3;"
Private Const DETAIL_CNDN = "^1|^4|^4|>4;"

Private Const FOOTER_RECEIPT1 = "<11|>3;"
Private Const FOOTER_RECEIPT1_1 = "<10|>2|>2;"
Private Const FOOTER_RECEIPT1_2 = "<8|>4|>2;"
Private Const FOOTER_RECEIPT1_3 = ">12|>2;"
Private Const FOOTER_RECEIPT2 = "<13|^3|>3;"
Private Const FOOTER_BILLS = "<13|^3|>3;"
Private Const FOOTER_CNDN = "<5|>4|>4;"

Private m_Enterprise As CEnterprise
Private m_Customer As CAPARMas
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As CName
Private m_Addr As CAddress
Private m_InventoryDoc As CInventoryDoc
Private m_ChildData As CInventoryDoc
Private m_OwnerAddress As CAddress

Private Sp As CSystemParam
Private X1 As Long
Private X2  As Long
Private Y1  As Long
Private Y2  As Long
Private AllEditEndPage As Boolean
Private m_InventoryDocIDColl As Collection
Private DriverIdCOll As Collection
Private TranSportorIdCOll As Collection
Private Mr1 As CMasterRef
Private Mr2 As CMasterRef
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_Addr = New CAddress
   Set m_OwnerAddress = New CAddress
   Set m_InventoryDocIDColl = New Collection
   Set DriverIdCOll = New Collection
   Set TranSportorIdCOll = New Collection
   Set m_Customer = New CAPARMas
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_Addr = Nothing
   Set m_OwnerAddress = Nothing
   Set m_InventoryDocIDColl = Nothing
   Set DriverIdCOll = Nothing
   Set TranSportorIdCOll = Nothing
   Set m_Customer = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
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

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("^10|^10|^10;", mdblWid, alngX)
   strPrint = ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   If Mr1.KEY_NAME <> "" Then
      strPrint = Mr1.KEY_NAME
      If Mr2.KEY_NAME <> "" Then
         strPrint = strPrint & "/" & Mr2.KEY_NAME
      End If
   ElseIf Mr2.KEY_NAME <> "" Then
         strPrint = Mr2.KEY_NAME
   Else
         strPrint = ".....................................;"
   End If
   
   strPrint = strPrint & "|..................................|.....................................;"
   Call Vsp.AddTable(strFormat, "", strPrint)
      
   strPrint = "ผู้ส่งสินค้า|ผู้รับสินค้า|ผู้จัดทำ;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
End Sub

Private Sub PrintBillHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempAddr As String
Dim Doc As CDocItem
Dim ItemCount As Long
Dim IsOK As Boolean
Dim EmpAddress As CEnterpriseAddress
Dim CustomerAddress As CApArAddress
Dim CusAddr As CAddress
Dim Emp As CEmployee

Dim Mr As CMasterRef

 Set TempRs = New ADODB.Recordset

   If I = 0 Then
      I = 1
      
      
      m_Enterprise.ENTERPRISE_ID = -1
      m_Enterprise.QueryFlag = 1
      If Not glbDaily.QueryEnterprise(m_Enterprise, TempRs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   Call m_Enterprise.PopulateFromRS(1, TempRs)
   m_Customer.QueryFlag = 1
   m_Customer.APAR_MAS_ID = m_InventoryDoc.GetFieldValue("APAR_MAS_ID")
   If Not glbDaily.QueryCustomer(m_Customer, TempRs, ItemCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call m_Customer.PopulateFromRS(1, TempRs)
      
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = m_Enterprise.ENTERPRISE_NAME & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Set EmpAddress = m_Enterprise.EnpAddresses(1)
   
   Set m_EnpAddress = EmpAddress.Addresses(1)
   
   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
   strPrint = "โทร. " & EmptyToString(m_EnpAddress.GetFieldValue("PHONE1"), "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.GetFieldValue("FAX1"), "N/A") & "|" & " ใบส่งเสริมการขาย " & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)

   strFormat = VSP_CalTable("<10;", mdblWid, alngX)
   strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & m_Enterprise.TAX_ID & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
   strPrint = "ลูกค้า| " & m_Customer.APAR_CODE & "|" & "เลขที่" & "ใบส่งเสริม" & "|" & m_InventoryDoc.GetFieldValue("DOCUMENT_NO") & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
   strPrint = m_Customer.APAR_NAME & "|" & "วันที่|" & DateToStringExtEx2((m_InventoryDoc.GetFieldValue("DOCUMENT_DATE"))) & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   If m_Customer.CstAddresses.Count > 0 Then
      Set CustomerAddress = m_Customer.CstAddresses(1)
      Set CusAddr = CustomerAddress.Addresses
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = EmptyToString(CusAddr.PackAddress, "N/A") & "||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   
   Set Emp = GetObject("CEmployee", m_EmployeeColl, Trim(Str(m_InventoryDoc.GetFieldValue("EMP_ID"))))
   
   strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
   strPrint = "หมายเหตุ   " & m_InventoryDoc.GetFieldValue("DOCUMENT_DESC") & "|ผู้รับผิดชอบ|" & Emp.EMP_NAME & " ( " & Emp.EMP_CODE & " )" & ";"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.FontBold = False
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
   
   Call PrintBillHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
      
   
   strFormat = VSP_CalTable(HEADER_RECEIPT1, mdblWid, alngX)
   strPrint = "No.|รายการ|จำนวน;"
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   X1 = Vsp.PageWidth - Vsp.MarginRight
   Y1 = Vsp.CurrentY
   
   
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
   Vsp.PaperSize = pprLetter
   
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
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
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Bi As CLotItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim DiscountAmount As Double
Dim VatAmount As Double
Dim IncludeDiscountAmount As Double
Dim IncludeVatAmount As Double
Dim TempVatAmount As Double
Dim TempDeposit As Double
Dim TempDiscount As Double
Dim TempCol As Collection
Dim SpaceCount As Long
Dim Ri As CRcpCnDn_Item
Dim Temp1 As Double
Dim Temp2 As Double
Dim TempInventoryDocID  As CInventoryDoc
Dim PageCount As Long

   RName = "genDoc"
   Set Rs = New ADODB.Recordset
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
   
   Call StartExportFile(Vsp)
   
   Set Rs = New ADODB.Recordset
   Set m_InventoryDoc = New CInventoryDoc
      
   Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", mcolParam("INVENTORY_DOC_ID"))
      
   m_InventoryDoc.QueryFlag = 1
   If Not glbDaily.QueryInventoryDoc(m_InventoryDoc, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
      
   Call m_InventoryDoc.PopulateFromRS(1, Rs)
   
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
      
      'Vsp.PenStyle = psDash
         
   Vsp.TableBorder = tbAll
      AllEditEndPage = True
      TempDiscount = 0
      TempVatAmount = 0
         
      Vsp.TableBorder = tbColumns
      Dim Addr As CAddress
'      If m_InventoryDoc.BRANCH_ADDRESS > 0 Then
'         Set Addr = New CAddress
'         Call Addr.SetFieldValue("ADDRESS_ID", m_InventoryDoc.BRANCH_ADDRESS)
'         Call Addr.QueryData(1, Rs, iCount)
'         If Not Rs.EOF Then
'            Call Addr.PopulateFromRS(1, Rs)
'            Call mcolParam.add(Addr.PackAddress, "BRANCH_ADDRESS")
'            Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "BRANCH_ADDRESS_PHONE")
'             Call mcolParam.add(Addr.GetFieldValue("FAX1"), "BRANCH_ADDRESS_FAX")
'            Call mcolParam.add(Addr.GetFieldValue("SHOW_LOCATION_FLAG"), "SHOW_LOCATION_FLAG")
'         End If
'         Set Addr = Nothing
'      Else
'         Call mcolParam.add("", "BRANCH_ADDRESS")
'         Call mcolParam.add("", "BRANCH_ADDRESS_PHONE")
'         Call mcolParam.add("", "BRANCH_ADDRESS_FAX")
'         Call mcolParam.add("", "SHOW_LOCATION_FLAG")
'      End If
      
      If PageCount > 1 Then
         Vsp.NewPage
      Else
         Vsp.StartDoc
         mblnHeader = True
      End If
        I = 0
      strFormat = VSP_CalTable(DETAIL_RECEIPT1, mdblWid, alngX)
      Set TempCol = m_InventoryDoc.ImportExportItems
      For Each Bi In TempCol
         I = I + 1
         strPrint = I & "|"
         strPrint = strPrint & Bi.PART_NO & " " & Bi.PART_DESC & "|"
         strPrint = strPrint & FormatNumber(MyDiff(Bi.TX_AMOUNT, Bi.UNIT_MULTIPLE)) & " " & Bi.UNIT_TRAN_NAME & ";"
         
               
         If I >= TempCol.Count Then
            Call AdjustPage(Vsp, strFormat, strPrint, 5000)
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      Next Bi

      Call PatchTable(Vsp, strFormat, "||||;", 2600) '10
         
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
         
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      AllEditEndPage = False
            
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
      Set Mr1 = GetObject("CMasterRef", DriverIdCOll, Trim(Str(m_InventoryDoc.GetFieldValue("DRIVER_ID"))))
      Set Mr2 = GetObject("CMasterRef", TranSportorIdCOll, Trim(Str(m_InventoryDoc.GetFieldValue("TRANSPORTOR_ID"))))
      
      Vsp.FontBold = False
      Call PrintBillFooter(alngX)
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
      If m_InventoryDoc.GetFieldValue("TRIP_FLAG") = "Y" Then
         Vsp = "*** จัดทริป ***"
      End If
      
       '--------------      Clear mColparam ----------------'
'   Call mcolParam.Remove("BRANCH_ADDRESS")
'   Call mcolParam.Remove("BRANCH_ADDRESS_PHONE")
'   Call mcolParam.Remove("BRANCH_ADDRESS_FAX")
'   Call mcolParam.Remove("SHOW_LOCATION_FLAG")
   '--------------      Clear mColparam ----------------'
               
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set m_InventoryDoc = Nothing
   
   genDoc = True
   
  'Vsp.PrintDoc (True)
   
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
   If AllEditEndPage Then
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
         
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
   End If
   
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
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   'Vsp.Paragraph = MODULE_DESC
   'Vsp.Paragraph = ""
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


