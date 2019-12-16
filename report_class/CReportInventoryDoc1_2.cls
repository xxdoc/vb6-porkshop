VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportInventoryDoc1_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportInventoryDoc1_2"
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
Private m_Bvis As Collection
Private m_BalanceLotColl As Collection
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
   Set m_Bvis = New Collection
   Set m_BalanceLotColl = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Bvis = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_BalanceLotColl = Nothing
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

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
      
   Call Cf.SetFieldValue(10, "^", "NO.", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "รหัสวัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(60, "^", "ชื่อวัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(10, "^", "คลัง", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(15, "^", "หน่วย", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "จ.คงคลัง", ">", , True)
   Call m_Headers1.AddField(Cf)
   If mcolParam("SHOW_LOT") = "1" Then
      Call Cf.SetFieldValue(20, "^", "จ.ตามLOT", ">", , True)
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "จ.ตามLOT", ">", , True)
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(20, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(20, "^", "ต.คงเหลือ", ">", , True)
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
   
   ToDate = "  ณ วันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.GetFieldValue("ENTERPRISE_NAME") & " " & AddStringFrontEnd(glbEnterPrise.GetFieldValue("BRANCH_NAME"), "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & ToDate & vbCrLf & _
                                       " จากสินค้า " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " ถึงสินค้า " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & " คลัง " & EmptyToString(mcolParam("LOCATION_NAME"), "N/A")
                                       
   Call GetBalanceByLotItemLink(m_Bvis, , mcolParam("FROM_DATE"), mcolParam("LOCATION_ID"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   Call MergeRecord
   
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
Dim m_LotItem As CLotItem
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
Dim Bvi As CLotItem
Dim HaveValue As Boolean
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

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
   Next j
   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      MasterInd = "20"
      Set m_LotItem = New CLotItem
      MasterInd = "1"
      
      m_LotItem.LOT_ITEM_ID = -1
      m_LotItem.TO_DOC_DATE = mcolParam("FROM_DATE")
      m_LotItem.LOCATION_ID = mcolParam("LOCATION_ID")
      m_LotItem.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
      m_LotItem.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
      m_LotItem.COUNT_AMOUNT = ""
      Call m_LotItem.QueryDataReport(20, Rs, iCount)
      
      I = 0
      HaveValue = False
      
      While Not Rs.EOF
         Call m_Details1.ClearField
         Call m_LotItem.PopulateFromRS(20, Rs)
         
         If PrevKey2 <> m_LotItem.PART_TYPE_NAME And I <> 0 Then
            If HaveValue Then
               Call GenerateFooter("", "รวม " & PrevKey1, "", Total1, m_Details1, 0)
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               HaveValue = False
            
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
            End If
         End If
         
         Set Bvi = GetObject("CLotItem", m_BalanceLotColl, Trim(m_LotItem.PART_ITEM_ID & "-" & m_LotItem.LOCATION_ID))
         
         If (Round(m_LotItem.SUM_AMOUNT, 2) <> 0) Or (Bvi.LOT_ITEM_AMOUNT - Bvi.TX_AMOUNT <> 0) Then
            PrevKey1 = m_LotItem.PART_TYPE_CODE
            PrevKey2 = m_LotItem.PART_TYPE_NAME
            
            I = I + 1
            
            Call m_Details1.ClearField
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = I
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = m_LotItem.PART_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = m_LotItem.PART_DESC
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(4)
            TempStr = m_LotItem.LOCATION_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
             
            Set HeadCf = m_Headers1.Fields(5)
            TempStr = m_LotItem.UNIT_CHANGE_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(6)
            TempStr = FormatNumberToNull(m_LotItem.SUM_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + m_LotItem.SUM_AMOUNT
            Total2(6) = Total2(6) + m_LotItem.SUM_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(7)
            TempStr = FormatNumberToNull(Bvi.LOT_ITEM_AMOUNT - Bvi.TX_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(7) = Total1(7) + Bvi.LOT_ITEM_AMOUNT - Bvi.TX_AMOUNT
            Total2(7) = Total2(7) + Bvi.LOT_ITEM_AMOUNT - Bvi.TX_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(8)
            If m_LotItem.CHK_STD_COST = "Y" Then
               TempStr = FormatNumberToNull(m_LotItem.COST_PER_AMOUNT)
            Else
               TempStr = FormatNumberToNull(Bvi.ACTUAL_UNIT_PRICE)
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(9)
            If m_LotItem.CHK_STD_COST = "Y" Then
               TempStr = FormatNumberToNull((m_LotItem.SUM_AMOUNT) * m_LotItem.COST_PER_AMOUNT)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(9) = Total1(9) + (m_LotItem.SUM_AMOUNT) * m_LotItem.COST_PER_AMOUNT
               Total2(9) = Total2(9) + (m_LotItem.SUM_AMOUNT) * m_LotItem.COST_PER_AMOUNT
            Else
               TempStr = FormatNumberToNull((m_LotItem.SUM_AMOUNT) * Bvi.ACTUAL_UNIT_PRICE)
               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
               Call m_Details1.AddField(BodyCf)
               Total1(9) = Total1(9) + (m_LotItem.SUM_AMOUNT) * Bvi.ACTUAL_UNIT_PRICE
               Total2(9) = Total2(9) + (m_LotItem.SUM_AMOUNT) * Bvi.ACTUAL_UNIT_PRICE
            End If
            
            If Not (HaveValue) Then
               Vsp = "ประเภทวัตถุดิบ : " & m_LotItem.PART_TYPE_NAME & "(" & m_LotItem.PART_TYPE_CODE & ")"
               HaveValue = True
            End If
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Rs.MoveNext
      Wend
      If HaveValue Then
         Call GenerateFooter("", "รวม " & PrevKey1, "", Total1, m_Details1, 0)           '
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Vsp.FontBold = True
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
      Call GenerateFooter("", "สุทธิ", "", Total2, m_Details1, 0)           '
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
            
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
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j))
      Else
         TempStr = ""
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
   Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.GetFieldValue("TAX_ID"))
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
Private Sub MergeRecord()
Dim LT As CLotItem
Dim TempLt  As CLotItem
   MasterInd = "21"
   For Each LT In m_Bvis
      Set TempLt = GetObject("CLotItem", m_BalanceLotColl, Trim(LT.PART_ITEM_ID & "-" & LT.LOCATION_ID), False)
      If TempLt Is Nothing Then
         Set TempLt = New CLotItem
         TempLt.PART_ITEM_ID = LT.PART_ITEM_ID
         TempLt.LOCATION_ID = LT.LOCATION_ID
         TempLt.PART_NO = LT.PART_NO
         TempLt.PART_DESC = LT.PART_DESC
         TempLt.LOT_ITEM_AMOUNT = LT.LOT_ITEM_AMOUNT
         TempLt.TX_AMOUNT = LT.TX_AMOUNT
         TempLt.ACTUAL_UNIT_PRICE = LT.ACTUAL_UNIT_PRICE
         TempLt.STD_COST = LT.STD_COST
         Call m_BalanceLotColl.add(TempLt, Trim(LT.PART_ITEM_ID & "-" & LT.LOCATION_ID))
      Else
         TempLt.LOT_ITEM_AMOUNT = TempLt.LOT_ITEM_AMOUNT + LT.LOT_ITEM_AMOUNT
         TempLt.TX_AMOUNT = TempLt.TX_AMOUNT + LT.TX_AMOUNT
         TempLt.ACTUAL_UNIT_PRICE = MyDiffEx(((TempLt.LOT_ITEM_AMOUNT - TempLt.TX_AMOUNT) * TempLt.ACTUAL_UNIT_PRICE) + ((LT.LOT_ITEM_AMOUNT - LT.TX_AMOUNT) * LT.ACTUAL_UNIT_PRICE), (TempLt.LOT_ITEM_AMOUNT - TempLt.TX_AMOUNT) + (LT.LOT_ITEM_AMOUNT - LT.TX_AMOUNT))
      End If
      Set TempLt = Nothing
   Next LT
   MasterInd = "1"
   Set m_Bvis = Nothing
End Sub
