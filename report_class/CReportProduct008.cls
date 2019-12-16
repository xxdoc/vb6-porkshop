VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportProduct008"
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

Private m_DistinctOutput1 As Collection
Private m_DistinctOutput2 As Collection
Private m_SumInputOutPutLost1 As Collection
Private m_SumInputOutPutLost2 As Collection

Private m_TagetJobs As Collection
Private m_TagetInputs As Collection
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
   Set m_DistinctOutput1 = New Collection
   Set m_DistinctOutput2 = New Collection
   Set m_SumInputOutPutLost1 = New Collection
   Set m_SumInputOutPutLost2 = New Collection
   
   Set m_TagetJobs = New Collection
   Set m_TagetInputs = New Collection
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
   Set m_DistinctOutput1 = Nothing
   Set m_DistinctOutput2 = Nothing
   
   Set m_SumInputOutPutLost1 = Nothing
   Set m_SumInputOutPutLost2 = Nothing
   
   Set m_TagetJobs = Nothing
   Set m_TagetInputs = Nothing
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

Private Sub printHeader(TypeShow As Long)
Dim strFormat As String
Dim strPrint As String
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim Ji As CJobItem
Dim TempProductionType As Long
Dim TempProductionTypeName   As String
Dim I As Long
Dim SumWidth As Double
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   
   Set Cf = New CReportField
   
   Call Cf.SetFieldValue(3, "^", "", ">")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(3, "^", "วัตถุดิบ", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(4, "^", "", ">")
   Call m_Headers1_1.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "ไข่รวม", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "TAGET", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "ACTUAL", ">", , True)
   Call m_Headers1.AddField(Cf)
      
   TempProductionType = -1
   I = 0
   For Each Ji In m_DistinctOutput1
      I = I + 1
      If Not (Ji.SUM_FLAG = "Y") Then
         If I <> 1 And (Ji.PRODUCTION_TYPE <> TempProductionType) Then
            Call Cf.SetFieldValue(8, "^", TempProductionTypeName, ">", , True)
            Call m_Headers0.AddField(Cf)
            SumWidth = SumWidth + 8
            Call Cf.SetFieldValue(2, "^", "TAGET", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "ACTUAL", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
         End If
         
         If Ji.PART_ITEM_ID > 0 Then
            If Not (TypeShow = 1) Then
               SumWidth = SumWidth + 2
               Call Cf.SetFieldValue(2, "^", Ji.BILL_DESC, ">", , True)
               Call m_Headers1.AddField(Cf)
            End If
         Else
            If Not (TypeShow = 1) Then
               SumWidth = SumWidth + 2
               Call Cf.SetFieldValue(2, "^", Ji.PROBLEM_SHORT_CODE, ">", , True)
               Call m_Headers1.AddField(Cf)
            End If
         End If
         
         TempProductionType = Ji.PRODUCTION_TYPE
         TempProductionTypeName = Ji.PRD_TYPE_SHORT_CODE
         
         If I = m_DistinctOutput1.Count Then
            Call Cf.SetFieldValue(8, "^", Ji.PRD_TYPE_SHORT_CODE, ">", , True)
            Call m_Headers0.AddField(Cf)
            
            SumWidth = SumWidth + 8
            
            Call Cf.SetFieldValue(2, "^", "TAGET", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "ACTUAL", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
            
         End If
         
      End If
      
   Next Ji
   
   Call Cf.SetFieldValue(SumWidth, "^", "แบ็ต 1 คัดล้าง", ">")
   Call m_Headers1_1.AddField(Cf)
            
   '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
   SumWidth = 0
   TempProductionType = -1
   I = 0
   For Each Ji In m_DistinctOutput2
      I = I + 1
      If Not (Ji.SUM_FLAG = "Y") Then
         If I <> 1 And (Ji.PRODUCTION_TYPE <> TempProductionType) Then
            Call Cf.SetFieldValue(8, "^", TempProductionTypeName, ">", , True)
            Call m_Headers0.AddField(Cf)
            
            SumWidth = SumWidth + 8
            
            Call Cf.SetFieldValue(2, "^", "TAGET", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "ACTUAL", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
         End If
         
         If Ji.PART_ITEM_ID > 0 Then
            If Not (TypeShow = 1) Then
               SumWidth = SumWidth + 2
               Call Cf.SetFieldValue(2, "^", Ji.BILL_DESC, ">", , True)
               Call m_Headers1.AddField(Cf)
            End If
         Else
            If Not (TypeShow = 1) Then
               SumWidth = SumWidth + 2
               Call Cf.SetFieldValue(2, "^", Ji.PROBLEM_SHORT_CODE, ">", , True)
               Call m_Headers1.AddField(Cf)
            End If
         End If
         
         TempProductionType = Ji.PRODUCTION_TYPE
         TempProductionTypeName = Ji.PRD_TYPE_SHORT_CODE
         
         If I = m_DistinctOutput2.Count Then
            Call Cf.SetFieldValue(8, "^", Ji.PRD_TYPE_SHORT_CODE, ">", , True)
            Call m_Headers0.AddField(Cf)
                        
            SumWidth = SumWidth + 8
            
            Call Cf.SetFieldValue(2, "^", "TAGET", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "ACTUAL", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
            
         End If
         
      End If
      
   Next Ji
   
   Call Cf.SetFieldValue(SumWidth, "^", "แบ็ต 2 บรรจุ", ">")
   Call m_Headers1_1.AddField(Cf)
   
   Set Cf = Nothing
            
   Call m_Headers1_1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
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
   
    FromDate = "จากวันที่ " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "ถึง " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = mcolParam("REPORT_NAME") & " " & FromDate & " " & ToDate & " สถานที่ผลิต " & EmptyToString(mcolParam("PRODUCTION_LOCATION_NAME"), "ทั้งหมด")
   
   'Batch 1
   Call GetDistinctJobOutItemByPrdType(m_DistinctOutput1, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), "(1)")
   Call GetSumJobInOutLostItem(m_SumInputOutPutLost1, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), "(1)")
   
   'Batch 2
   Call GetDistinctJobOutItemByPrdType(m_DistinctOutput2, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), "(2)")
   Call GetSumJobInOutLostItem(m_SumInputOutPutLost2, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("PRODUCTION_LOCATION"), "(2)")
                                                            
   Call LoadTagetJobByType(m_TagetJobs, mcolParam("MONTH_ID"), Val(mcolParam("YEAR_NO") - 543))
   Call LoadTagetJobInputByType(m_TagetInputs, mcolParam("MONTH_ID"), Val(mcolParam("YEAR_NO") - 543))
                                                                  
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
Dim m_JobItem As CJobItem
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100)  As Double

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim Ji  As CJobItem
Dim TempJi  As CJobItem

Dim TjInputAmount As Double
Dim TjOutput1Amount As Double
Dim TjOutput2Amount As Double
Dim InputAmount As Double

Dim TempProductionType   As Long
Dim TempProductionTypeName  As String
Dim SumProductionType As Double

Dim m_TagetJobDetails As CTagetJobDetail
Dim m_Inputs As CTagetJob
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

      Set m_JobItem = New CJobItem
      
      m_JobItem.JOB_ID = -1
      m_JobItem.FROM_DATE = mcolParam("FROM_BILL_DATE")
      m_JobItem.TO_DATE = mcolParam("TO_BILL_DATE")
      m_JobItem.PRD_LOCATION_ID = mcolParam("PRODUCTION_LOCATION")
      m_JobItem.BATCH_NO_SET = "(1,2)"
      Call m_JobItem.QueryData(23, Rs, iCount)
      
      I = 0
      
      If Vsp.CurrentPage = 1 Then
         Call printHeader(1)
      End If
      
      While Not Rs.EOF
         I = I + 1
         
         Call m_Details1.ClearField
         Call m_JobItem.PopulateFromRS(23, Rs)
         
         Call m_Details1.ClearField
         
         k = 0
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = m_JobItem.BILL_DESC
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set m_Inputs = GetObject("CTagetJob", m_TagetInputs, Trim(Str(m_JobItem.PART_ITEM_ID)))
         
         TjInputAmount = m_Inputs.INPUT_AMOUNT
         
         'TAGETรับไข่
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(TjInputAmount, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + TjInputAmount
         Total2(k) = Total2(k) + TjInputAmount
         
         InputAmount = m_JobItem.TX_AMOUNT
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(InputAmount, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + InputAmount
         Total2(k) = Total2(k) + InputAmount
         
         TempProductionType = -1
         SumProductionType = 0
         j = 0
         For Each Ji In m_DistinctOutput1
            If Not (Ji.SUM_FLAG = "Y") Then
               j = j + 1
               
               If (j <> 1 And (Ji.PRODUCTION_TYPE <> TempProductionType)) Then
                  Set m_TagetJobDetails = GetObject("CTagetJobDetail", m_TagetJobs, Trim("1-" & m_JobItem.PART_ITEM_ID & "-" & TempProductionType))
                  'Getobject จาก TagetJob
                  k = k + 1
                  TjOutput1Amount = MyDiff(TjInputAmount * m_TagetJobDetails.OUTPUT_AMOUNT, 100)
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(TjOutput1Amount, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + TjOutput1Amount
                  Total2(k) = Total2(k) + TjOutput1Amount
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(m_TagetJobDetails.OUTPUT_AMOUNT)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(SumProductionType, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + SumProductionType
                  Total2(k) = Total2(k) + SumProductionType
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff(SumProductionType * 100, m_JobItem.TX_AMOUNT))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  SumProductionType = 0
               End If
               
               TempProductionType = Ji.PRODUCTION_TYPE
               TempProductionTypeName = Ji.PRD_TYPE_SHORT_CODE
               
               'k = k + 1
               Set TempJi = GetObject("CJobItem", m_SumInputOutPutLost1, Trim("1" & "-" & m_JobItem.PART_ITEM_ID & "-" & Ji.PART_ITEM_ID & "-" & Ji.LOST_ID), True)
'               If Not (TypeShow = 1) Then
'                  Set HeadCf = m_Headers1.Fields(k)
'                  TempStr = FormatNumberToNull(TempJi.TX_AMOUNT, 0)
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details1.AddField(BodyCf)
'               End If
'               Total1(k) = Total1(k) + TempJi.TX_AMOUNT
'               Total2(k) = Total2(k) + TempJi.TX_AMOUNT
               SumProductionType = SumProductionType + TempJi.TX_AMOUNT
               
               If (j = m_DistinctOutput1.Count) Then
                  'Getobject จาก TagetJob
                  Set m_TagetJobDetails = GetObject("CTagetJobDetail", m_TagetJobs, Trim("1-" & m_JobItem.PART_ITEM_ID & "-" & TempProductionType))
                  k = k + 1
                  TjOutput1Amount = MyDiff(TjInputAmount * m_TagetJobDetails.OUTPUT_AMOUNT, 100)
                  
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(TjOutput1Amount, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + TjOutput1Amount
                  Total2(k) = Total2(k) + TjOutput1Amount
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(m_TagetJobDetails.OUTPUT_AMOUNT)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(SumProductionType, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + SumProductionType
                  Total2(k) = Total2(k) + SumProductionType
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff(SumProductionType * 100, m_JobItem.TX_AMOUNT))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  SumProductionType = 0
               End If
               
            End If
            
         Next Ji
         
         TempProductionType = -1
         SumProductionType = 0
         j = 0
         For Each Ji In m_DistinctOutput2
            If Not (Ji.SUM_FLAG = "Y") Then
               j = j + 1
               
               If (j <> 1 And (Ji.PRODUCTION_TYPE <> TempProductionType)) Then
                  'Getobject จาก TagetJob
                  Set m_TagetJobDetails = GetObject("CTagetJobDetail", m_TagetJobs, Trim("2-" & m_JobItem.PART_ITEM_ID & "-" & TempProductionType))
                  k = k + 1
                  TjOutput2Amount = MyDiff(TjInputAmount * m_TagetJobDetails.OUTPUT_AMOUNT, 100)
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(TjOutput2Amount, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + TjOutput2Amount
                  Total2(k) = Total2(k) + TjOutput2Amount
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(m_TagetJobDetails.OUTPUT_AMOUNT)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(SumProductionType, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + SumProductionType
                  Total2(k) = Total2(k) + SumProductionType
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff(SumProductionType * 100, m_JobItem.TX_AMOUNT))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  SumProductionType = 0
               End If
               
               TempProductionType = Ji.PRODUCTION_TYPE
               TempProductionTypeName = Ji.PRD_TYPE_SHORT_CODE
               
 '              k = k + 1
               Set TempJi = GetObject("CJobItem", m_SumInputOutPutLost2, Trim("2" & "-" & m_JobItem.PART_ITEM_ID & "-" & Ji.PART_ITEM_ID & "-" & Ji.LOST_ID), True)
'               If Not (TypeShow = 1) Then
'                  Set HeadCf = m_Headers1.Fields(k)
'                  TempStr = FormatNumberToNull(TempJi.TX_AMOUNT, 0)
'                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'                  Call m_Details1.AddField(BodyCf)
'               End If
'               Total1(k) = Total1(k) + TempJi.TX_AMOUNT
'               Total2(k) = Total2(k) + TempJi.TX_AMOUNT
               SumProductionType = SumProductionType + TempJi.TX_AMOUNT
               
               If (j = m_DistinctOutput2.Count) Then
                  'Getobject จาก TagetJob
                  Set m_TagetJobDetails = GetObject("CTagetJobDetail", m_TagetJobs, Trim("2-" & m_JobItem.PART_ITEM_ID & "-" & TempProductionType))
                  k = k + 1
                  TjOutput2Amount = MyDiff(TjInputAmount * m_TagetJobDetails.OUTPUT_AMOUNT, 100)
                  
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(TjOutput2Amount, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + TjOutput2Amount
                  Total2(k) = Total2(k) + TjOutput2Amount
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(m_TagetJobDetails.OUTPUT_AMOUNT)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(SumProductionType, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + SumProductionType
                  Total2(k) = Total2(k) + SumProductionType
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff(SumProductionType * 100, m_JobItem.TX_AMOUNT))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  SumProductionType = 0
               End If
               
            End If
            
         Next Ji
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         
         Rs.MoveNext
      Wend
         
      Call GenerateFooter("สุทธิ", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Vsp = ""
      
      '------------------------------------------------------------------------------------------------------------------------------------------------------------
      
      
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
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j), 0)
      ElseIf j > 4 Then
         If j Mod 4 = 1 Then
            TempStr = FormatNumberToNull(MyDiff(Tot(j - 1) * 100, Tot(2)))
         ElseIf j Mod 4 = 3 Then
            TempStr = FormatNumberToNull(MyDiff(Tot(j - 1) * 100, Tot(3)))
         End If
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
Dim strFormat   As String
Dim strPrint  As String
Dim alngX() As Long
Dim tbdt As TableBorderSettings
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
   
   tbdt = Vsp.TableBorder
   
   Vsp.TableBorder = tbdt
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
   'Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.TAX_ID)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taLeftMiddle
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   Vsp.FontSize = iSize
   If mblnHeader Then
'      If Vsp.CurrentPage > 1 Then
'         Call printHeader
'      End If
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
Private Sub GenerateHeader(txt1 As String, Details As CFieldList)
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
         Call BodyCf.SetFieldValue(HeadCf.Width, "<", TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub
