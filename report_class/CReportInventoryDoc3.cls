VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventoryDoc3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventoryDoc3"
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

Private InventoryBals2 As Collection
Private MovementPartType As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Details1 = New CFieldList
   
   Set Rs = New ADODB.Recordset
   Set InventoryBals2 = New Collection
   Set MovementPartType = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set InventoryBals2 = Nothing
   Set MovementPartType = Nothing
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
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField

   Call Cf.SetFieldValue(9, "^", "รายละเอียดเอกสาร", "^")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(12, "^", "รับเข้า", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(12, "^", "จ่ายออก", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คงเหลือ", ">")
   Call m_Headers0.AddField(Cf)
   If mcolParam("SHOW_UNIT_NAME") = 1 Then
      Call Cf.SetFieldValue(2, "^", "หน่วย", "<")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "หน่วย", "<")
      Call m_Headers0.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(1, "^", "No", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รหัส", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "รายการ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ยกมา", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับเข้า", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ซื้อ", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ผลิต", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "รับคืน", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอน", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปรับยอด", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เบิกใช้", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ผลิต", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "โอน", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ปรับยอด", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ขาย", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "คืน", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "จำนวน", ">", , True)
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_UNIT_NAME") = 1 Then
      Call Cf.SetFieldValue(2, "^", "", ">")
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "", ">")
      Call m_Headers1.AddField(Cf)
   End If
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
                                       " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " คลัง " & EmptyToString(mcolParam("LOCATION_NAME"), "คลังรวม")
                                       
   Call GetSumMovementPartItemType(MovementPartType, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("LOCATION_ID"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
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
Dim NewDate As Date

Dim m_StockCode As CStockCode
Dim Ba As CLotItem
Dim TempLi As CLotItem
Dim SumRow As Double

Dim PrevKey1 As String
Dim PrevKey2 As String

Dim HaveValue As Boolean
Dim HaveHeader As Boolean

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

   Set InventoryBals2 = New Collection
   NewDate = DateAdd("D", -1, mcolParam("FROM_BILL_DATE"))
   If mcolParam("FROM_BILL_DATE") > 0 Then
      Call LoadLeftAmount(InventoryBals2, -1, NewDate, mcolParam("LOCATION_ID"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   End If
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   
   I = 0
   Set BodyCf = New CReportField
   
   Set m_StockCode = New CStockCode
   
   m_StockCode.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   m_StockCode.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   m_StockCode.EXPENSE_FLAG = "N"
   Call m_StockCode.QueryData(4, Rs, iCount)
   
   HaveHeader = False
   
   While (Not Rs.EOF)
      
      Call m_StockCode.PopulateFromRS(4, Rs)
      
      If PrevKey1 <> Trim(m_StockCode.STOCK_TYPE_CODE) Then
         If HaveHeader Then
            Call GenerateFooter("รวม", "", "", Total1, m_Details1)
            Call m_Details1.GetString(2, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            HaveHeader = False
         End If
      End If
      
      PrevKey1 = Trim(m_StockCode.STOCK_TYPE_CODE)
      
      Call m_Details1.ClearField
      I = I + 1
      HaveValue = False
      SumRow = 0
      For j = 1 To m_Headers1.Fields.Count
         Set HeadCf = m_Headers1.Fields(j)
         
         Amt = 0
         If j = 1 Then
            TempStr = I
         ElseIf j = 2 Then
            TempStr = m_StockCode.STOCK_NO
         ElseIf j = 3 Then
            TempStr = m_StockCode.STOCK_DESC
         ElseIf j = 4 Then
            Set Ba = GetObject("CLotItem", InventoryBals2, Trim(Str(m_StockCode.STOCK_CODE_ID)))
            Amt = Ba.SUM_AMOUNT
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 5 Then
            ' รับเข้า
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim(IMPORT_DOCTYPE & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 6 Then
            ' ใบเสร็จซื้อสด
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("22" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = TempLi.TX_AMOUNT
            ' ใบรับสินค้าซื้อ
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("11" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = Amt + TempLi.TX_AMOUNT
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 7 Then
            ' ใบผลิต
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("1000" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = Amt + TempLi.TX_AMOUNT
            
'            If (m_StockCode.STOCK_NO = "41-0003") Then
'               Amt = 832
'            ElseIf (m_StockCode.STOCK_NO = "42-3001") Then
'               Amt = 39486
'            ElseIf (m_StockCode.STOCK_NO = "42-1001") Then
'               Amt = 46330
'            End If
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 8 Then
            ' รับคืน
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("30" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = Amt + TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 9 Then
            ' โอน
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("3" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = Amt + TempLi.TX_AMOUNT
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 10 Then
            ' ปรับยอด
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("4" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = Amt + TempLi.TX_AMOUNT
            'ตั้งยอด
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("5" & "-" & m_StockCode.STOCK_CODE_ID & "-I"))
            Amt = Amt + TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow + Amt
         ElseIf j = 11 Then
            ' เบิกใช้
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("2" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow - Amt
         ElseIf j = 12 Then
            ' ผลิต
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("1000" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow - Amt
         ElseIf j = 13 Then
            ' โอน
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("3" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            
'            If (m_StockCode.STOCK_NO = "42-3001") Then
'               Amt = 0
'            ElseIf (m_StockCode.STOCK_NO = "42-1001") Then
'               Amt = 11300
'            End If
            
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow - Amt
         ElseIf j = 14 Then
            ' ปรับยอด
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("4" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            'ตั้งยอด
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("5" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow - Amt
         ElseIf j = 15 Then
            ' ใบขายสด
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("21" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            ' ใบขายเชื่อ
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("10" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow - Amt
         ElseIf j = 16 Then
            ' ใบคืน
            Set TempLi = GetObject("CLotItem", MovementPartType, Trim("31" & "-" & m_StockCode.STOCK_CODE_ID & "-E"))
            Amt = Amt + TempLi.TX_AMOUNT
            TempStr = FormatNumberToNull(Amt, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + Amt
            Total2(j) = Total2(j) + Amt
            SumRow = SumRow - Amt
         ElseIf j = 17 Then
            ' คงเหลือ
            TempStr = FormatNumberToNull(SumRow, Val(mcolParam("DECIMAL_AMOUNT")))
            Total1(j) = Total1(j) + SumRow
            Total2(j) = Total2(j) + SumRow
         ElseIf j = 18 Then
            TempStr = m_StockCode.UNIT_CHANGE_NAME
         End If
         
         HaveValue = CheckHaveValue(HaveValue, Amt)
         
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
         Call m_Details1.AddField(BodyCf)
      Next j
      
      If HaveValue Then
         If Not (HaveHeader) Then
            Vsp = m_StockCode.STOCK_TYPE_NAME
            HaveHeader = True
         End If
         Call m_Details1.GetString(2, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      Else
         I = I - 1
      End If
      
      Rs.MoveNext
   Wend
   '=== รายการ
   
   Call GenerateFooter("รวม", "", "", Total1, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("=", "", "", Total2, m_Details1)
   Call m_Details1.GetString(2, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Set InventoryBals2 = Nothing
   Set m_StockCode = Nothing
   
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
      ElseIf j = 2 Then
         TempStr = txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf HeadCf.SumFlag Then
         If (j = m_Headers1.Fields.Count) Or (j = m_Headers1.Fields.Count - 1) Then
            TempStr = FormatNumberToNull(Tot(j), Val(mcolParam("DECIMAL_AMOUNT")))
         Else
            TempStr = FormatNumberToNull(Tot(j), Val(mcolParam("DECIMAL_AMOUNT")))
         End If
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
