VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling039"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling039"
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
Private m_CountAllCycleDriver As Collection
Private m_CountAllCycleTransportor As Collection
Private m_SumAmountByDriver As Collection
Private m_SumAmountByTranSportor As Collection
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
   Set m_CountAllCycleDriver = New Collection
   Set m_CountAllCycleTransportor = New Collection
   Set m_SumAmountByDriver = New Collection
   Set m_SumAmountByTranSportor = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
   Set m_Details2 = Nothing
   Set m_CountAllCycleDriver = Nothing
   Set m_CountAllCycleTransportor = Nothing
   Set m_SumAmountByDriver = Nothing
   Set m_SumAmountByTranSportor = Nothing
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
   
   Call Cf.SetFieldValue(4, "^", "เส้นทางขนส่ง", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "พนักงาน", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ประเภทรถ", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "เที่ยว", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "จุด", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ยอดฟอง", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "ค่าขนส่ง/เที่ยว", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "รวมจ่าย", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "เฉลี่ย/ฟอง", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "หมายเหตุ", "<")
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
Dim DocTypeSet As String
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
   
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
                                             
   If (mcolParam("PO_FLAG") = 1) Then
      MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " (ใบ PO) "
                                       
      DocTypeSet = "(" & PO_DOCTYPE & ")"
   ElseIf mcolParam("INVOICE_FLAG") = 1 Then
      MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " (ใบส่งของ และ ขายสด) "
      DocTypeSet = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")"
   Else
      MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " (ใบ PO) "
      DocTypeSet = "(" & PO_DOCTYPE & ")"
   End If
                                                   
   Call GetCountAllCycleDriver(m_CountAllCycleDriver, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), DocTypeSet)
   Call GetCountAllCycleTranSportor(m_CountAllCycleTransportor, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), DocTypeSet)
                                             
   Call GetSumAmountByDriver(m_SumAmountByDriver, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), DocTypeSet, mcolParam("INCLUDE_FREE"))
   Call GetSumAmountByTranSportor(m_SumAmountByTranSportor, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), DocTypeSet, mcolParam("INCLUDE_FREE"))
                                          
                                          
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
Dim m_TranSport As CTranSportDetail
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
Dim TempBd As CBillingDoc
Dim TempBd1 As CBillingDoc
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

   For j = 1 To UBound(Total2)
      Total2(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
      I = 0
      Set BodyCf = New CReportField

      Set m_TranSport = New CTranSportDetail
      
      m_TranSport.TRANSPORT_DETAIL_ID = -1
      m_TranSport.ORDER_BY = 4
      Call m_TranSport.QueryData(1, Rs, iCount)
      
      I = 0
      
      While Not Rs.EOF
         I = I + 1
         Call m_TranSport.PopulateFromRS(1, Rs)
         
         Call m_Details1.ClearField
         
         If Len(m_TranSport.TRANSPORTOR_NAME) > 0 Then
            PrevKey2 = m_TranSport.TRANSPORTOR_NAME
         Else
            PrevKey2 = m_TranSport.DRIVER_NAME
         End If
         
         If PrevKey1 <> PrevKey2 Then
            k = 0
            k = k + 1
            
            If Len(m_TranSport.TRANSPORTOR_NAME) > 0 Then
               PrevKey1 = m_TranSport.TRANSPORTOR_NAME
            Else
               PrevKey1 = m_TranSport.DRIVER_NAME
            End If
            
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = m_TranSport.TRANSPORT_PATH
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            If Len(m_TranSport.TRANSPORTOR_NAME) > 0 Then
               TempStr = m_TranSport.TRANSPORTOR_NAME
            Else
               TempStr = m_TranSport.DRIVER_NAME
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = m_TranSport.CAR_TYPE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Len(m_TranSport.TRANSPORTOR_NAME) > 0 Then
               Set TempBd = GetObject("CBillingDoc", m_CountAllCycleTransportor, Trim(Str(m_TranSport.TRANSPORTOR_ID)))
            Else
               Set TempBd = GetObject("CBillingDoc", m_CountAllCycleDriver, Trim(Str(m_TranSport.DRIVER_ID)))
            End If
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = TempBd.TRANSPORT_CYCLE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = "" 'จุด
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If Len(m_TranSport.TRANSPORTOR_NAME) > 0 Then
               Set TempBd1 = GetObject("CBillingDoc", m_SumAmountByTranSportor, Trim(Str(m_TranSport.TRANSPORTOR_ID)))
            Else
               Set TempBd1 = GetObject("CBillingDoc", m_SumAmountByDriver, Trim(Str(m_TranSport.DRIVER_ID)))
            End If
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(TempBd1.TOTAL_AMOUNT, 0)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = FormatNumberToNull(m_TranSport.COST_PER_ROUND)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            If m_TranSport.COST_PER_ROUND > 0 Then
               TempStr = FormatNumberToNull(TempBd.TRANSPORT_CYCLE * m_TranSport.COST_PER_ROUND) 'รวมค่าใช้จ่าย
            Else
               TempStr = FormatNumberToNull(Getexpense(m_TranSport.TRANSPORTOR_ID))
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            If m_TranSport.COST_PER_ROUND > 0 Then
               TempStr = FormatNumberToNull(MyDiffEx(TempBd.TRANSPORT_CYCLE * m_TranSport.COST_PER_ROUND, TempBd1.TOTAL_AMOUNT), 3) 'ค่าเฉลี่ย/ฟอง
            Else
               TempStr = FormatNumberToNull(MyDiffEx(Getexpense(m_TranSport.TRANSPORTOR_ID), TempBd1.TOTAL_AMOUNT), 3) 'ค่าเฉลี่ย/ฟอง
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = "" 'หมายเหตุ
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            If TempBd1.TOTAL_AMOUNT > 0 Then
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
         End If
         Rs.MoveNext
      Wend
      
'       Call GenerateFooter("รวม", "", "", Total1, m_Details1, 0)
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
'
'      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
'      Call m_Details1.GetString(1, TempStr1, TempStr2)
'      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'      Vsp.FontBold = True
'      Call Vsp.AddTable(strFormat, "", TempStr2)
'      Vsp.FontBold = False
      
      Call GenerateEndDoc
      
      Vsp.TableBorder = tbNone
      Call PatchTable(Vsp, strFormat, ";", 1000) '10
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      TempStr2 = "................................................|................................................;"
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      TempStr2 = "ผู้จัดทำ|ผู้ตรวจสอบ;"
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
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
'      ElseIf j = 3 Then
'         TempStr = FormatNumberToNull(Tot(j))
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
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*รวมรายการของแถม"
   End If
End Sub
Private Function Getexpense(TranSportorID As Long) As Double
   If TranSportorID = mcolParam("TRANSPORTOR1") Then
      Getexpense = mcolParam("TRANSPORTOR1_PRICE")
      Exit Function
   End If
   If TranSportorID = mcolParam("TRANSPORTOR2") Then
      Getexpense = mcolParam("TRANSPORTOR2_PRICE")
      Exit Function
   End If
   If TranSportorID = mcolParam("TRANSPORTOR3") Then
      Getexpense = mcolParam("TRANSPORTOR3_PRICE")
      Exit Function
   End If
End Function
