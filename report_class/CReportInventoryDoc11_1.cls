VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventoryDoc11_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventoryDoc5"
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
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 10

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Details1 As CFieldList

Private MovementPartType As Collection
Private DistincePartOutput As Collection
Private DistincePartInput As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   
   Set Rs = New ADODB.Recordset
   Set MovementPartType = New Collection
   Set DistincePartOutput = New Collection
   Set DistincePartInput = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set MovementPartType = Nothing
   Set DistincePartOutput = Nothing
   Set DistincePartInput = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
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
Dim TempStr2 As String
Dim Cf As CReportField
Dim TempLt  As CLotItem
Dim SumWidth As Long
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(8, "^", "รายละเอียด", "^")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "วันที่", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "เลขที่", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2 * DistincePartInput.Count, "^", "วัตถุดิบ", "^")
   Call m_Headers0.AddField(Cf)
   
   For Each TempLt In DistincePartInput
      Call Cf.SetFieldValue(2, "^", TempLt.PART_DESC, ">", , True)
      Call m_Headers1.AddField(Cf)
   Next TempLt
   
   SumWidth = 0
   For Each TempLt In DistincePartOutput
      Call Cf.SetFieldValue(2, "^", TempLt.PART_DESC, ">", , True)
      Call m_Headers1.AddField(Cf)
      
      If CheckInPercent(TempLt.PART_NO) Then
         Call Cf.SetFieldValue(2, "^", "%", ">")
         Call m_Headers1.AddField(Cf)
         SumWidth = SumWidth + 1
      End If
   Next TempLt
   Call Cf.SetFieldValue(2, "^", "รวม", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2 * (DistincePartOutput.Count + 1 + SumWidth), "^", "สินค้า", "^")
   Call m_Headers0.AddField(Cf)
   
   Set Cf = Nothing
            
   Call m_Headers0.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr2
   Call Vsp.AddTable(strFormat, "", strPrint)
            
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
   
   FromDate = " จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = " ถึงวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & FromDate & ToDate & vbCrLf & _
                                       " จากวัตถุดิบ " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงวัตถุดิบ " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " จากคลัง " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
   '
   Call GetDistinctPartOutputByInput(DistincePartOutput, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("LOCATION_ID"), 1000, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   Call GetDistinctPartInputByInput(DistincePartInput, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("LOCATION_ID"), 1000, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
   Call GetSumAmountByInputOutput(MovementPartType, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("LOCATION_ID"), 1000)
   
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim Amt As Double

Dim Li As CLotItem
Dim TempLi As CLotItem
Dim TempDataLi As CLotItem
Dim SumRow As Double
Dim SumInput As Double
Dim PrevKey1 As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
'
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j

   Set Li = New CLotItem
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   '=== รายการ
   Set Li = New CLotItem
   
   Li.LOT_ITEM_ID = -1
   Li.FROM_DOC_DATE = mcolParam("FROM_BILL_DATE")
   Li.TO_DOC_DATE = mcolParam("TO_BILL_DATE")
   Li.LOCATION_ID = mcolParam("LOCATION_ID")
   Li.ORDER_BY = mcolParam("ORDER_BY")
   Li.ORDER_TYPE = mcolParam("ORDER_TYPE")
   Li.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   Li.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   Li.DOCUMENT_TYPE = 1000

'   Li.INVENTORY_SUB_TYPE = mcolParam("INVENTORY_SUB_TYPE")
   Call Li.QueryData(45, Rs, iCount, False)
   
   While (Not Rs.EOF)
      
      Call Li.PopulateFromRS(45, Rs)
         
      Call m_Details1.ClearField
      I = I + 1
      
      SumRow = 0
      SumInput = 0
      
      j = 0
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = DateToStringExtEx2(Li.DOCUMENT_DATE)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = Li.DOCUMENT_NO
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
      
      For Each TempLi In DistincePartInput
         j = j + 1
         Set TempDataLi = GetObject("CLotItem", MovementPartType, Trim(Li.INVENTORY_DOC_ID & "-" & TempLi.PART_ITEM_ID & "-E"))
         Amt = TempDataLi.TX_AMOUNT
         TempStr = FormatNumberToNull(Amt, 0)
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Total1(j) = Total1(j) + Amt
         Total2(j) = Total2(j) + Amt
         SumInput = SumInput + Amt
      Next TempLi
      
      For Each TempLi In DistincePartOutput
         j = j + 1
         Set TempDataLi = GetObject("CLotItem", MovementPartType, Trim(Li.INVENTORY_DOC_ID & "-" & TempLi.PART_ITEM_ID & "-I"))
         Amt = TempDataLi.TX_AMOUNT
         TempStr = FormatNumberToNull(Amt, 0)
         Set HeadCf = m_Headers1.Fields(j)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
         
         Total1(j) = Total1(j) + Amt
         Total2(j) = Total2(j) + Amt
         SumRow = SumRow + Amt
         
         If CheckInPercent(TempLi.PART_NO) Then
            j = j + 1
            Amt = MyDiffEx(TempDataLi.TX_AMOUNT * 100, SumInput)
            TempStr = FormatNumberToNull(Amt, 2)
            Set HeadCf = m_Headers1.Fields(j)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
            Call m_Details1.AddField(BodyCf)
         End If
      Next TempLi
      
      j = j + 1
      Set HeadCf = m_Headers1.Fields(j)
      TempStr = FormatNumberToNull(SumRow, 0)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call m_Details1.AddField(BodyCf)
      Total1(j) = Total1(j) + SumRow
      Total2(j) = Total2(j) + SumRow
      
      Call m_Details1.GetString(2, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
      Rs.MoveNext
   Wend
   '=== รายการ
   
   Call GenerateFooter("สุทธิ", "", "", Total1, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set Li = Nothing
   
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.Description
End Function

Private Sub GenerateFooter(txt1 As String, txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList)
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
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j), 0)
      ElseIf (Not HeadCf.SumFlag) And j > 3 Then
         TempStr = FormatNumberToNull(MyDiffEx(Tot(j - 1) * 100, Tot(m_Headers1.Fields.Count)), 2)
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
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
Private Function CheckInPercent(StockNo) As Boolean
Dim Status As Boolean
   Status = False
   If Not (mcolParam("FROM_STOCK_NO_PERCENT") = "") Then
      If (StockNo >= mcolParam("FROM_STOCK_NO_PERCENT")) Then
         Status = True
      Else
         Status = False
         CheckInPercent = Status
         Exit Function
      End If
   End If
   If Not (mcolParam("TO_STOCK_NO_PERCENT") = "") Then
      If (StockNo <= mcolParam("TO_STOCK_NO_PERCENT")) Then
         Status = True
      Else
         Status = False
         CheckInPercent = Status
         Exit Function
      End If
   End If
   CheckInPercent = Status
End Function


