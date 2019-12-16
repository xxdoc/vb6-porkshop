VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CReportBilling003
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling005"
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
Private Const DETAIL_SIZE = 12

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

Private m_PartItem As Collection
Private m_Sum As Collection
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
   
   Set m_PartItem = New Collection
   Set m_Sum = New Collection
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
   
   Set m_PartItem = Nothing
   Set m_Sum = Nothing
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
Dim Part As CPrintLabel
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = False
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   
   Set Cf = New CReportField
   
'   Call Cf.SetFieldValue(1, "^", "N", "^")
'   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(30, "^", "สาขา", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "รหัส", "<")
   Call m_Headers1.AddField(Cf)
   
   For Each Part In m_PartItem
      Call Cf.SetFieldValue(15, "^", Part.LABEL_DESC, ">")
      Call m_Headers1.AddField(Cf)
      
      Call Cf.SetFieldValue(20, "^", "ราคา", ">")
      Call m_Headers1.AddField(Cf)
   Next Part
   
   Call Cf.SetFieldValue(20, "^", "จำนวน", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(25, "^", "ยอดเงิน", ">")
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
Dim FromDate As String
Dim ToDate As String
   
   mstrErrMsg = ""
   mblnHeader = False
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 500
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " จากลูกค้า " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " ถึงลูกค้า " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A")
                                       
   Call LoadDistinctLabelPartItemEx(m_PartItem, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("INCLUDE_FREE"))
   Call LoadSumBranchPartItem(m_Sum, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("INCLUDE_FREE"))
   
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
Dim m_PrintLabel As CPrintLabel
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Sc As CStockCode
Dim PartItems As Collection
Dim strPrint As String
Dim Part As CPrintLabel
Dim SumTotal As CPrintLabel
Dim SumRow1 As Double
Dim SumRow2 As Double

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
   Set Sc = New CStockCode
   Set PartItems = New Collection
   'Call LoadStockCode(Sc, Nothing, PartItems)
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
   Next j
   For j = 1 To UBound(Total2)
      Total2(j) = 0
   Next j
   
   
      I = 0
      Set BodyCf = New CReportField

      Set m_PrintLabel = New CPrintLabel
      
      Call m_PrintLabel.SetFieldValue("PRINT_LABEL_ID", -1)
      Call m_PrintLabel.SetFieldValue("FROM_DATE", mcolParam("FROM_BILL_DATE"))
      Call m_PrintLabel.SetFieldValue("TO_DATE", mcolParam("TO_BILL_DATE"))
      
      m_PrintLabel.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
      m_PrintLabel.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
      m_PrintLabel.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
      m_PrintLabel.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
      m_PrintLabel.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
      m_PrintLabel.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
      m_PrintLabel.FREE_FLAG = StringToFreeFlag(mcolParam("INCLUDE_FREE"))
      
       Call m_PrintLabel.QueryData(9, Rs, iCount)
      
      I = 0
      
      PrevKey1 = ""
      
      While Not Rs.EOF
         I = I + 1
         Call m_PrintLabel.PopulateFromRS(9, Rs)
         
         SumRow1 = 0
         SumRow2 = 0
         
         If PrevKey1 <> Getkey(m_PrintLabel) Then
            If I <> 1 Then
               Call GenerateFooter("", "รวม", "", Total2, m_Details1, 0)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = False
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               For j = 1 To UBound(Total2)
                  Total2(j) = 0
               Next j
               
               I = 1
               mblnHeader = False
               Vsp.NewPage
            Else
               Call StartExportFile(Vsp)
               Vsp.StartDoc
            End If
            PrevKey1 = Getkey(m_PrintLabel)
            Vsp.FontBold = False
            Vsp.TableBorder = tbNone
            strFormat = VSP_CalTable("<4;", mdblWid, alngX)
            strPrint = m_PrintLabel.GetFieldValue("SALE_LONG_NAME") & " " & m_PrintLabel.GetFieldValue("SALE_LAST_NAME") & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
            Vsp = ""
            Vsp.TableBorder = tbAll
            Call printHeader
            mblnHeader = True
         End If
         
         Call m_Details1.ClearField
         
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = m_PrintLabel.GetFieldValue("BRANCH_NAME")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = m_PrintLabel.GetFieldValue("BRANCH_CODE")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         j = 2
         For Each Part In m_PartItem
            j = j + 1
            Amt = 0
            Set SumTotal = GetObject("CPrintLabel", m_Sum, Trim(m_PrintLabel.GetFieldValue("EMP_ID") & "-" & m_PrintLabel.GetFieldValue("BRANCH_ID") & "-" & Part.PART_ITEM_ID))
            Amt = Amt + SumTotal.TOTAL_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(Amt, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Total2(j) = Total2(j) + Amt
            SumRow1 = SumRow1 + Amt
            
            j = j + 1
            
            Amt = 0
            Set SumTotal = GetObject("CPrintLabel", m_Sum, Trim(m_PrintLabel.GetFieldValue("EMP_ID") & "-" & m_PrintLabel.GetFieldValue("BRANCH_ID") & "-" & Part.PART_ITEM_ID))
            Amt = Amt + SumTotal.TOTAL_PRICE - SumTotal.DISCOUNT_AMOUNT - SumTotal.EXTRA_DISCOUNT_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(j)
            TempStr = FormatNumberToNull(Amt)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            SumRow2 = SumRow2 + Amt
            Total2(j) = Total2(j) + Amt
            
         Next Part
         
         j = j + 1
         
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(SumRow1, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Total2(j) = Total2(j) + SumRow1
         
         j = j + 1
         
         Set HeadCf = m_Headers1.Fields(j)
         TempStr = FormatNumberToNull(SumRow2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Total2(j) = Total2(j) + SumRow2
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Rs.MoveNext
      Wend
      
      Call GenerateFooter("", "รวม", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = False
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      Vsp = ""
      
      Call GenerateEndDoc
      
      Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Set m_PrintLabel = Nothing
   
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
   Vsp.FontBold = False
   
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
Private Function Getkey(Pl As CPrintLabel) As String
      Getkey = Trim(Str(Pl.GetFieldValue("EMP_ID")))
End Function
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub

