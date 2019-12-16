VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportMain004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements CReportInterface

Private Const MODULE_NAME = "CReportMain004"
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
Private Sp As CSystemParam
Private m_ReportConfig As CReportConfig

Private m_cData As CBillingDoc

Private m_ImageWidth As Double
Private m_ImageHeight As Double

Private CollAparAddress As Collection
Private CollEnpAddress As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_ReportConfig = New CReportConfig
   Set CollAparAddress = New Collection
   Set CollEnpAddress = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_ReportConfig = Nothing
   Set CollAparAddress = Nothing
   Set CollEnpAddress = Nothing
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

'Private Sub PrintPOFooter(alngX() As Long, BD As CBillingDoc)
'      Call GenerateDocFooter(2, Vsp, mcolParam, tbNone, BD)
'End Sub
Private Sub PrintPOHeader(alngX() As Long)
Dim TempBorder As TableBorderSettings
Dim TempMarginLeft As Double
'
   TempMarginLeft = Vsp.MarginLeft
   Vsp.MarginLeft = (2.5 * 567) + m_ReportConfig.GetFieldValue("DUMMY_OFFSET")
'
   Call GenerateDocHeader(m_ReportConfig.GetFieldValue("HEAD_OFFSET"), tbNone)
'
'
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
   Vsp.MarginBottom = "0cm"
   Vsp.MarginFooter = "0cm"
   Vsp.MarginHeader = "1cm"
   Vsp.MarginTop = "0cm"
   Vsp.FontName = "AngsanaUPC"
   
   If Val(mcolParam("FONT_SIZE")) > 0 Then
      Vsp.FontSize = Val(mcolParam("FONT_SIZE"))
   Else
      Vsp.FontSize = DETAIL_SIZE
   End If
   'MODULE_DESC = MapText("ใบส่งสินค้า")
   Call LoadApArMasAddressLetter(CollAparAddress, mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), mcolParam("APAR_TYPE"), mcolParam("APAR_GROUP"))
   
   Dim Addr As CAddress
   Set Addr = New CAddress
   Call LoadEnterpriseAddress(Addr, Nothing, CollEnpAddress, False)
   
   
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
Dim J As Long
Dim TempBorder As TableBorderSettings
Dim Addr As CAddress
   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   Call StartExportFile(Vsp)
   
   Vsp.StartDoc
   
   TempBorder = tbNone
   
   Call GenerateDocBody(TempBorder)
   
'   Call PrintPOFooter(alngX, m_cData)
      
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   
   genDoc = True
      
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
      
   'Vsp.DrawPicture mcolParam("BACK_GROUND"), "0.0cm", "0.4cm", "20.4cm", "27.8cm", 10
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

'Public Sub GenerateDocFooter(Ind As Long, Vsp As VSPrinter, mcolParam As Collection, TempBorder As TableBorderSettings, BD As CBillingDoc)
'Dim TempX As Long
'Dim TempY As Long
'Dim Amt As Double
'Dim Sum1 As Double
'Dim Sum2 As Double
'Dim Sum3 As Double
'Dim TempFont As Byte
'Dim TempString As String
'Dim Mr  As CMasterRef
'
'   Vsp.FontBold = True
'   Vsp.TableBorder = TempBorder
'
'   TempX = Vsp.CurrentX
'   TempY = Vsp.CurrentY
'   TempFont = Vsp.FontSize
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 5
'   Vsp.TableCell(tcRows) = 1
'    Vsp.TableCell(tcRowHeight) = "2.56cm"
'    Vsp.TableCell(tcColWidth, , 1) = "2.5cm"
'    Vsp.TableCell(tcColWidth, , 2) = "8cm"
'    Vsp.TableCell(tcColWidth, , 3) = "4.5cm"
'    Vsp.TableCell(tcColWidth, , 4) = "1.3cm"
'    Vsp.TableCell(tcColWidth, , 5) = "3cm"
'
'   Vsp.FontSize = 12
'   If mcolParam("SHOW_LOCATION_FLAG") = "Y" Then
'      Vsp.TableCell(tcText, 1, 1) = "สถานที่ส่งของ"
'      Vsp.TableCell(tcAlign, 1, 1) = taLeftTop
'
'      'Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME") & TempString
'      Vsp.TableCell(tcText, 1, 2) = mcolParam("CUSTOMER_NAME") & " (สาขา " & BD.CUSTOMER_BRANCH_NAME & ")" & vbCrLf & mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX")
'      Vsp.TableCell(tcAlign, 1, 2) = taLeftTop
'   End If
'   Vsp.TableBorder = TempBorder
'   Vsp.EndTable
'
'
'   Vsp.CurrentX = TempX
'   Vsp.CurrentY = TempY
'   Vsp.FontSize = TempFont
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 5
'   Vsp.TableCell(tcRows) = 5
'    Vsp.TableCell(tcRowHeight) = "0.64cm"
'    Vsp.TableCell(tcColWidth, , 1) = "2.5cm"
'    Vsp.TableCell(tcColWidth, , 2) = "8cm"
'    Vsp.TableCell(tcColWidth, , 3) = "4.5cm"
'    Vsp.TableCell(tcColWidth, , 4) = "1.3cm"
'    Vsp.TableCell(tcColWidth, , 5) = "3cm"
'
'   Vsp.TableCell(tcAlign, 1, 5) = taRightMiddle
'   Vsp.TableCell(tcText, 1, 5) = FormatNumber(BD.TOTAL_PRICE)
'
'   Sum1 = Val(FormatNumber(BD.DISCOUNT_AMOUNT, , False)) + Val(FormatNumber(BD.EXT_DISCOUNT_AMOUNT, , False))
'
'   If BD.EXT_DISCOUNT_PERCENT > 0 Then
'      Vsp.TableCell(tcAlign, 2, 4) = taCenterMiddle
'      Vsp.TableCell(tcText, 2, 4) = BD.EXT_DISCOUNT_PERCENT & " %"
'   End If
'
'   Vsp.TableCell(tcAlign, 2, 5) = taRightMiddle
'   Vsp.TableCell(tcText, 2, 5) = FormatNumber(Sum1)
'
'   Sum2 = Val(FormatNumber(BD.TOTAL_PRICE, , False)) - Sum1
'
'   Vsp.TableCell(tcAlign, 3, 5) = taRightMiddle
'   Vsp.TableCell(tcText, 3, 5) = FormatNumber(Sum2)
'
'   If BD.VAT_PERCENT > 0 Then
'      Vsp.TableCell(tcAlign, 4, 4) = taCenterMiddle
'      Vsp.TableCell(tcText, 4, 4) = BD.VAT_PERCENT & " %"
'   End If
'
'   Vsp.TableCell(tcAlign, 4, 5) = taRightMiddle
'   Vsp.TableCell(tcText, 4, 5) = FormatNumber(BD.VAT_AMOUNT)
'
'   Amt = Sum2 + Val(FormatNumber(BD.VAT_AMOUNT, , False))
'
'   Vsp.TableCell(tcAlign, 5, 2) = taRightMiddle
'   Vsp.TableCell(tcText, 5, 2) = "(" & ThaiBaht(Amt) & ")"
'
'   Vsp.TableCell(tcAlign, 5, 5) = taRightMiddle
'   Vsp.TableCell(tcText, 5, 5) = FormatNumber(Amt)
'
'   Vsp.TableBorder = TempBorder
'   Vsp.EndTable
'
'   '   '============
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 1
'   Vsp.TableCell(tcRows) = 1
'    Vsp.TableCell(tcRowHeight, 1) = 2.3 * 567
'    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
'    Vsp.TableBorder = TempBorder
'    Vsp.EndTable
'
'      '   '============
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 3
'   Vsp.TableCell(tcRows) = 1
'   Vsp.TableCell(tcRowHeight, 1) = 1 * 567
'   Vsp.TableCell(tcColWidth, 1, 1) = "9.3cm"
'   Vsp.TableCell(tcColWidth, 1, 2) = "3.9cm"
'   Vsp.TableCell(tcColWidth, 1, 3) = "6.2cm"
'
'   'Set Mr = GetObject("CMasterRef", CollAparAddress, Trim(Str(BD.DRIVER_ID)), False)
'   If Not (Mr Is Nothing) Then
'      Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'      Vsp.TableCell(tcText, 1, 2) = Mr.KEY_NAME
'   Else
'      'Set Mr = GetObject("CMasterRef", TranSportorIdCOll, Trim(Str(BD.TRANSPORTOR_ID)), False)
'      If Not (Mr Is Nothing) Then
'         Vsp.TableCell(tcAlign, 1, 2) = taCenterMiddle
'         Vsp.TableCell(tcText, 1, 2) = Mr.KEY_NAME
'      End If
'    End If
'
'    Vsp.TableBorder = TempBorder
'    Vsp.EndTable
'
'   '   '============
'   Vsp.StartTable
'   Vsp.TableCell(tcCols) = 1
'   Vsp.TableCell(tcRows) = 1
'    Vsp.TableCell(tcRowHeight, 1) = 1.2 * 567
'    Vsp.TableCell(tcColWidth, 1, 1) = "19.3cm"
'    Vsp.TableBorder = TempBorder
'
'    Vsp.EndTable
''   '============
'
'End Sub
Public Sub GenerateDocHeader(HeaderOffset As Double, TempBorder As TableBorderSettings)
   
   '============
   Vsp.StartTable
   Vsp.TableCell(tcCols) = 1
   Vsp.TableCell(tcRows) = 1
   Vsp.TableCell(tcRowHeight, 1) = (1.5 * 567) + Val(HeaderOffset)
    Vsp.TableCell(tcColWidth, 1, 1) = "17.5cm"
    Vsp.TableBorder = TempBorder
    Vsp.EndTable
   '============
End Sub

Public Sub GenerateDocBody(TempBorder As TableBorderSettings)
Dim TempApArAddress As CAPARMas
Dim I As Long
Dim J As Long
Dim Amount As Long
Dim TempCountAmount As Long
Dim TempAddr As CAddress

   Amount = Val(mcolParam("AMOUNT_PER_ITEM"))
   If Amount <= 0 Then
      Amount = 1
   End If
   I = 0
   J = 0
   If Val(mcolParam("AMOUNT_ENTERPRISE")) > 0 Then
      For Each TempAddr In CollEnpAddress
         For TempCountAmount = 1 To Val(mcolParam("AMOUNT_ENTERPRISE"))
            I = I + 1
            J = J + 1
      
            If J > 27 Then
               J = 1
               Vsp.NewPage
            End If
            
            If J Mod 3 = 1 Then
               Vsp.StartTable
                Vsp.TableCell(tcCols) = 5
               Vsp.TableCell(tcRows) = 1
                Vsp.TableCell(tcRowHeight) = "3.05cm"
               Vsp.TableBorder = TempBorder
         
                Vsp.TableCell(tcColWidth, , 1) = "5.5cm"
                Vsp.TableCell(tcColWidth, , 2) = "0.5cm"
                Vsp.TableCell(tcColWidth, , 3) = "5.5cm"
                Vsp.TableCell(tcColWidth, , 4) = "0.5cm"
                Vsp.TableCell(tcColWidth, , 5) = "5.5cm"
         
         
                Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
                Vsp.TableCell(tcText, 1, 1) = glbEnterPrise.ENTERPRISE_NAME & vbCrLf & TempAddr.PackAddress
             ElseIf J Mod 3 = 2 Then
               Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
               Vsp.TableCell(tcText, 1, 3) = glbEnterPrise.ENTERPRISE_NAME & vbCrLf & TempAddr.PackAddress
            ElseIf J Mod 3 = 0 Then
               Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
               Vsp.TableCell(tcText, 1, 5) = glbEnterPrise.ENTERPRISE_NAME & vbCrLf & TempAddr.PackAddress
               
               Vsp.EndTable
            End If
         Next TempCountAmount
      Next TempAddr
   End If
   
   For Each TempApArAddress In CollAparAddress
      For TempCountAmount = 1 To Amount
         I = I + 1
         J = J + 1
   
         If J > 27 Then
            J = 1
            Vsp.NewPage
         End If
         
         If J Mod 3 = 1 Then
            Vsp.StartTable
             Vsp.TableCell(tcCols) = 5
            Vsp.TableCell(tcRows) = 1
             Vsp.TableCell(tcRowHeight) = "3.05cm"
            Vsp.TableBorder = TempBorder
      
             Vsp.TableCell(tcColWidth, , 1) = "5.5cm"
             Vsp.TableCell(tcColWidth, , 2) = "0.5cm"
             Vsp.TableCell(tcColWidth, , 3) = "5.5cm"
             Vsp.TableCell(tcColWidth, , 4) = "0.5cm"
             Vsp.TableCell(tcColWidth, , 5) = "5.5cm"
      
      
             Vsp.TableCell(tcAlign, 1, 1) = taLeftMiddle
             Vsp.TableCell(tcText, 1, 1) = TempApArAddress.APAR_NAME & vbCrLf & TempApArAddress.APAR_ADDRESS
          ElseIf J Mod 3 = 2 Then
            Vsp.TableCell(tcAlign, 1, 3) = taLeftMiddle
            Vsp.TableCell(tcText, 1, 3) = TempApArAddress.APAR_NAME & vbCrLf & TempApArAddress.APAR_ADDRESS
         ElseIf J Mod 3 = 0 Then
            Vsp.TableCell(tcAlign, 1, 5) = taLeftMiddle
            Vsp.TableCell(tcText, 1, 5) = TempApArAddress.APAR_NAME & vbCrLf & TempApArAddress.APAR_ADDRESS
            
            Vsp.EndTable
         End If
      Next TempCountAmount
   Next TempApArAddress
   
   If J Mod 3 > 0 Then
      Vsp.EndTable
   End If
End Sub
