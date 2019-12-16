VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportPrintLabel003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements CReportInterface

Private Const MODULE_NAME = "CReportPrintLabel003"
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
Private m_ReportConfig As CReportConfig
Private Rs As ADODB.Recordset

Private Const TITLE_SIZE = 32
Private Const HEADER_SIZE = 32
Private Const DETAIL_SIZE = 32

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

Private Sub Class_Initialize()
   Set m_ReportConfig = New CReportConfig
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
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   Set m_ReportConfig = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
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
   
'   Vsp.FontSize = HEADER_SIZE
   Vsp.FontSize = 18
   Vsp.FontBold = True
   Vsp.TableBorder = tbNone
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(10, "<", "", "<")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(10, "<", "", "<")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers0.AddField(Cf)
      
   Call Cf.SetFieldValue(10, "<", "", "<")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(10, "^", "", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(10, "^", "", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(2, "^", "", "^")
   Call m_Headers1.AddField(Cf)
      
   Call Cf.SetFieldValue(10, "^", "", "^")
   Call m_Headers1.AddField(Cf)

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
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 50
   Vsp.MarginFooter = 50
   Vsp.MarginHeader = 50
   Vsp.MarginLeft = 793.7007
   Vsp.MarginRight = 793.7007
   Vsp.MarginTop = 398.4
   Vsp.FontName = "AngsanaUPC"
'   Vsp.FontSize = DETAIL_SIZE
   Vsp.FontSize = 18
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"), m_ReportConfig, False)

   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   
   Vsp.MarginLeft = Vsp.MarginLeft + m_ReportConfig.GetFieldValue("DUMMY_OFFSET")
   
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
Dim iCount As Long
Dim TempStr As String
Dim TempColl As Collection
Dim TempX As Double
Dim TempY As Double
Dim LbColl As Collection
Dim TempLb As CPrintLabel
Dim CountRow As Long
Dim TempLine As Double

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set LbColl = New Collection
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   I = 0
   Set BodyCf = New CReportField

   Set m_PrintLabel = New CPrintLabel
      
   Call m_PrintLabel.SetFieldValue("BILLING_DOC_ID", mcolParam("BILLING_DOC_ID"))
   Call m_PrintLabel.QueryData(2, Rs, iCount)
   
   While Not Rs.EOF
      Call m_PrintLabel.PopulateFromRS(2, Rs)
      Set TempLb = GetObject("CPrintLable", LbColl, m_PrintLabel.GetFieldValue("BLOCK_ID") & "-" & m_PrintLabel.GetFieldValue("BRANCH_ID"), False)
      If TempLb Is Nothing Then
         Set TempLb = New CPrintLabel
         Call TempLb.SetFieldValue("TOTAL_AMOUNT", m_PrintLabel.GetFieldValue("PACK_AMOUNT"))
         Call TempLb.SetFieldValue("BLOCK_ID", m_PrintLabel.GetFieldValue("BLOCK_ID"))
         Call TempLb.SetFieldValue("BRANCH_ID", m_PrintLabel.GetFieldValue("BRANCH_ID"))
         Call TempLb.SetFieldValue("PART_DESC", m_PrintLabel.GetFieldValue("PART_DESC"))
         Call LbColl.add(TempLb, TempLb.GetFieldValue("BLOCK_ID") & "-" & TempLb.GetFieldValue("BRANCH_ID"))
         Set TempLb = Nothing
      Else
         Call TempLb.SetFieldValue("TOTAL_AMOUNT", TempLb.GetFieldValue("TOTAL_AMOUNT") + m_PrintLabel.GetFieldValue("PACK_AMOUNT"))
      End If
      Rs.MoveNext
   Wend
   
   I = 0
   CountRow = 0
   TempLine = 0
   Vsp.TableBorder = tbNone
   If Not Rs.BOF Then
      Rs.MoveFirst
   End If
   Vsp.CurrentY = 398.4 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")     'ตำแหน่งบนสุดที่พิมพ์ขอบเส้น
   While Not Rs.EOF
      Call m_PrintLabel.PopulateFromRS(2, Rs)
      For j = 1 To m_PrintLabel.GetFieldValue("PACK_AMOUNT")
         I = I + 1
         If I = 1 Then
            Call m_Details1.ClearField
         End If

         Set TempLb = GetObject("CPrintLable", LbColl, m_PrintLabel.GetFieldValue("BLOCK_ID") & "-" & m_PrintLabel.GetFieldValue("BRANCH_ID"))
         
         If mcolParam("MODE_TYPE") = 2 Then    'ตรวจ BLOCK ของ BIG C
            Set HeadCf = m_Headers0.Fields(I)
            TempStr = "                            ใบกำกับสินค้า" & vbCrLf & "      รายการสินค้า " & m_PrintLabel.GetFieldValue("PART_DESC") & vbCrLf & "      แผนก...PERISHABLE...(560)" & vbCrLf & "      สาขา " & m_PrintLabel.GetFieldValue("BRANCH_CODE") & vbCrLf & "      ชื่อร้านค้า บริษัท ควอลิตี้ มีท จำกัด" & vbCrLf & "      วันที่ " & mcolParam("TEMP_DUE_DATE")
            Vsp.FontSize = 10
            Vsp.LineSpacing = 75
         Else
            Set HeadCf = m_Headers1.Fields(I)
            TempStr = m_PrintLabel.GetFieldValue("BLOCK_NAME") & "   " & m_PrintLabel.GetFieldValue("BRANCH_CODE") & vbCrLf & m_PrintLabel.GetFieldValue("LABEL_DESC") & "   " & (1 + TempLb.GetFieldValue("ITEM_AMOUNT")) & " / " & TempLb.GetFieldValue("TOTAL_AMOUNT")
            Vsp.LineSpacing = 120
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Call TempLb.SetFieldValue("ITEM_AMOUNT", 1 + TempLb.GetFieldValue("ITEM_AMOUNT"))
         If I Mod 5 = 0 Then 'เมื่อครบ 5 แถวก็ทำการพิมพ์ลงไป 1 บรรทัด
            If CountRow = 9 Then 'ครบ 9 บรรทัดให้ขึ้นหน้าใหม่
               Vsp.NewPage
               CountRow = 0
               Vsp.CurrentY = 398.4 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            End If
            TempLine = Vsp.CurrentY
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            I = 0
            CountRow = CountRow + 1
'            Vsp.CurrentY = Vsp.CurrentY
            If CountRow = 1 Then
               Vsp.CurrentY = 2267.716536 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 2 Then
               Vsp.CurrentY = 4138.582677 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 3 Then
               Vsp.CurrentY = 6009.4488 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 4 Then
               Vsp.CurrentY = 7880.314961 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 5 Then
               Vsp.CurrentY = 9751.181103 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 6 Then
               Vsp.CurrentY = 11622.04724 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 7 Then
               Vsp.CurrentY = 13549.6063 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            ElseIf CountRow = 8 Then
               Vsp.CurrentY = 15363.77953 + m_ReportConfig.GetFieldValue("HEAD_OFFSET")
            End If
         Else
            I = I + 1
            If mcolParam("MODE_TYPE") = 2 Then
               Set HeadCf = m_Headers0.Fields(I)
            Else
               Set HeadCf = m_Headers1.Fields(I)
            End If
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
      Next j
      Rs.MoveNext
   Wend

   If I <> 0 Then
      For I = (I + 1) To 5
         If mcolParam("MODE_TYPE") = 2 Then
            Set HeadCf = m_Headers0.Fields(I)
            Vsp.FontSize = 10
            Vsp.LineSpacing = 75
         Else
            Set HeadCf = m_Headers1.Fields(I)
            Vsp.SpaceBefore = 156
            Vsp.SpaceAfter = 156
         End If
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
      Next I
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Call Vsp.AddTable(strFormat, "", TempStr2)
      I = 0
   End If
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set LbColl = Nothing
   Set TempLb = Nothing
   
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
      ElseIf j = 7 Then
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
   
'   Vsp.FontSize = 8
'   dY(0) = Vsp.CurrentY
'   Vsp.TextAlign = taLeftBottom
'   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
'   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExtEx2(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.TAX_ID)
'   Vsp.CurrentY = dY(0)
   
'   Vsp.FontSize = TITLE_SIZE
'   Vsp.FontBold = True
'   Vsp.TextAlign = taCenterBottom
'   Vsp.Paragraph = MODULE_DESC
'   Vsp.Paragraph = ""
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
