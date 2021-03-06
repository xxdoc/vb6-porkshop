VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportProduct007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportJobItemSumItem"
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
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 8

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
Private m_DistinctOutput As Collection
Private m_SumInputOutPutLostMain As Collection
Private m_SumInputOutPutLostNoMain As Collection
Private m_SumInputM As Collection
Private m_SumInputMNoMain As Collection

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
   Set m_DistinctOutput = New Collection
   Set m_SumInputOutPutLostMain = New Collection
   Set m_SumInputOutPutLostNoMain = New Collection
   Set m_SumInputM = New Collection
   Set m_SumInputMNoMain = New Collection
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
   Set m_DistinctOutput = Nothing
   Set m_SumInputOutPutLostMain = Nothing
   Set m_SumInputOutPutLostNoMain = Nothing
   Set m_SumInputM = Nothing
   Set m_SumInputMNoMain = Nothing
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
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim HeadCf As CReportField
Dim Ji As CJobItem
Dim TempProductionType As Long
Dim TempProductionTypeName   As String
Dim I As Long
Dim j  As Long

   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   If mcolParam("DAIRY_REPORT") = 1 Then
      Call Cf.SetFieldValue(3, "^", "�ѵ�شԺ", ">")
   Call m_Headers0.AddField(Cf)
   End If
   Call Cf.SetFieldValue(3, "^", "Ƿ LOT", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��", ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�.�", ">")
   Call m_Headers0.AddField(Cf)
   
   If mcolParam("DAIRY_REPORT") = 1 Then
      Call Cf.SetFieldValue(3, "^", "", "<")
      Call m_Headers1.AddField(Cf)
   End If
   Call Cf.SetFieldValue(3, "^", "", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   TempProductionType = -1
   I = 0
   j = 0
   For Each Ji In m_DistinctOutput
      I = I + 1
      If Not (Ji.SUM_FLAG = "Y") Then
         If I <> 1 And (Ji.PRODUCTION_TYPE <> TempProductionType) Then
            j = j + 1
            Call Cf.SetFieldValue(4 * j, "^", TempProductionTypeName, ">", , True)
            Call m_Headers0.AddField(Cf)
            j = 0
            Call Cf.SetFieldValue(2, "^", "���", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
         End If
         
         If CheckHaveCollumnDetail(Ji) Then
            j = j + 1
            If Ji.PART_ITEM_ID > 0 Then
               Call Cf.SetFieldValue(2, "^", Ji.BILL_DESC, ">", , True)
               Call m_Headers1.AddField(Cf)
               Call Cf.SetFieldValue(2, "^", "%", ">")
               Call m_Headers1.AddField(Cf)
            Else
               Call Cf.SetFieldValue(2, "^", Ji.PROBLEM_SHORT_CODE, ">", , True)
               Call m_Headers1.AddField(Cf)
               Call Cf.SetFieldValue(2, "^", "%", ">")
               Call m_Headers1.AddField(Cf)
            End If
         End If
         TempProductionType = Ji.PRODUCTION_TYPE
         TempProductionTypeName = Ji.PRD_TYPE_SHORT_CODE
         
         If I = m_DistinctOutput.Count Then
            j = j + 1
            Call Cf.SetFieldValue(4 * j, "^", TempProductionTypeName, ">", , True)
            Call m_Headers0.AddField(Cf)
            
            Call Cf.SetFieldValue(2, "^", "���", ">", , True)
            Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "%", ">")
            Call m_Headers1.AddField(Cf)
         End If
         
      End If
      
   Next Ji
   
   Set Cf = Nothing
            
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
Dim FDate As Date
Dim TDate As Date
   
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
   
    FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = mcolParam("REPORT_NAME") & " " & FromDate & " " & ToDate & " �ҡ�ѵ�شԺ " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " �֧ " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A") & vbCrLf & " �ҡ��Ե�ѳ�� " & EmptyToString(mcolParam("FROM_STOCK_NO1"), "N/A") & " �֧ " & EmptyToString(mcolParam("TO_STOCK_NO1"), "N/A") & " BATCH : " & EmptyToString(mcolParam("BATCH_NO_SET"), "������")
   
   FDate = DateAdd("D", -1 * mcolParam("SUB_FROM_DATE"), mcolParam("FROM_BILL_DATE"))       ' -1 ��͹������ ������������ѧ 1 ��͹
   TDate = DateAdd("D", mcolParam("ADD_TO_DATE"), mcolParam("TO_BILL_DATE"))                ' 1 ��͹������ �����Ź�˹�� 1 ��͹
   Call GetDistinctJobOutItemByPrdTypeEx(m_DistinctOutput, FDate, TDate, , GetBatchNoSet(), mcolParam("FROM_STOCK_NO1"), mcolParam("TO_STOCK_NO1"))
   
   Call GetSumJobInOutLostItemByInPutDocDate(m_SumInputOutPutLostMain, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), , GetBatchNoSet(), True, , , , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_STOCK_NO1"), mcolParam("TO_STOCK_NO1"))
   Call GetSumJobInOutLostItemByInPutDocDate(m_SumInputOutPutLostNoMain, FDate, TDate, , GetBatchNoSet(), , , , True, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_STOCK_NO1"), mcolParam("TO_STOCK_NO1"))
   
   Call GetJobInputByPartItemInPutDate(m_SumInputM, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), , , True, , , , mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   Call GetJobInputByPartItemInPutDate(m_SumInputMNoMain, FDate, TDate, , , , , , True, mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"))
   
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
Dim JiEx  As CJobItem
Dim TempJi  As CJobItem
Dim TempJiEx  As CJobItem
Dim InProcessAmount As Double
Dim TempProductionType   As Long
Dim TempProductionTypeName  As String
Dim SumProductionType As Double
Dim ProductAmount As Double
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
   Next j
   Call StartExportFile(Vsp)
   Vsp.StartDoc
      I = 0
      Set BodyCf = New CReportField

      Set m_JobItem = New CJobItem
      
      m_JobItem.JOB_ID = -1
      m_JobItem.FROM_DATE = mcolParam("FROM_BILL_DATE")
      m_JobItem.TO_DATE = mcolParam("TO_BILL_DATE")
      m_JobItem.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
      m_JobItem.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
      m_JobItem.MainLocationPrdFlag = True
      m_JobItem.BATCH_NO_SET = GetBatchNoSet()
      Call m_JobItem.QueryData(12, Rs, iCount)
      
      I = 0
      
      While Not Rs.EOF
         I = I + 1
         
         If I = 1 And mcolParam("PRINT_TO_FILE") = 1 Then
            Call printHeader
         End If
         
         Call m_Details1.ClearField
         Call m_JobItem.PopulateFromRS(12, Rs)
         
         If PrevKey1 <> Trim(Str(m_JobItem.PART_ITEM_ID)) And I > 1 Then
            If mcolParam("DAIRY_REPORT") = 0 Then
               If mcolParam("SHOW_SUMMARY") > 0 Then
                  Call GenerateFooter(PrevKey2, "", "", Total1, m_Details1, 0)
               Else
                  Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
                  Vsp.FontBold = True
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               
               If mcolParam("SHOW_SUMMARY") <= 0 Then
                  Call GenerateHeader("�ѵ�شԺ  :  " & m_JobItem.PART_DESC & " (" & m_JobItem.PART_NO & ")", m_Details1)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               End If
               For j = 1 To UBound(Total1)
                  Total1(j) = 0
               Next j
            End If
         ElseIf I = 1 Then
            If mcolParam("DAIRY_REPORT") = 0 Then
               If mcolParam("SHOW_SUMMARY") <= 0 Then
                  Call GenerateHeader("�ѵ�شԺ  :  " & m_JobItem.PART_DESC & " (" & m_JobItem.PART_NO & ")", m_Details1)
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
               End If
            End If
         End If
         
         Call m_Details1.ClearField
         PrevKey1 = Trim(Str(m_JobItem.PART_ITEM_ID))
         PrevKey2 = m_JobItem.BILL_DESC
         
         k = 0
         If mcolParam("DAIRY_REPORT") = 1 Then
            k = k + 1
            Set HeadCf = m_Headers1.Fields(k)
            TempStr = m_JobItem.BILL_DESC
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         End If
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = DateToStringExtEx2(m_JobItem.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(m_JobItem.WEIGHT_AMOUNT, 2)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         k = k + 1
         Set HeadCf = m_Headers1.Fields(k)
         TempStr = FormatNumberToNull(m_JobItem.TX_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + m_JobItem.TX_AMOUNT
         Total2(k) = Total2(k) + m_JobItem.TX_AMOUNT
         
         k = k + 1
         Set Ji = GetObject("CJobItem", m_SumInputM, Trim(m_JobItem.PART_ITEM_ID & "-" & m_JobItem.DOCUMENT_DATE & "-" & m_JobItem.WEIGHT_AMOUNT))
         Set JiEx = GetObject("CJobItem", m_SumInputMNoMain, Trim(m_JobItem.PART_ITEM_ID & "-" & m_JobItem.DOCUMENT_DATE & "-" & m_JobItem.WEIGHT_AMOUNT))
         Set HeadCf = m_Headers1.Fields(k)
         ProductAmount = Ji.TX_AMOUNT + JiEx.TX_AMOUNT
         TempStr = FormatNumberToNull(Ji.TX_AMOUNT + JiEx.TX_AMOUNT, 0)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(k) = Total1(k) + Ji.TX_AMOUNT + JiEx.TX_AMOUNT
         Total2(k) = Total2(k) + Ji.TX_AMOUNT + JiEx.TX_AMOUNT
         
         TempProductionType = -1
         SumProductionType = 0
         j = 0
         For Each Ji In m_DistinctOutput
            If Not (Ji.SUM_FLAG = "Y") Then
               j = j + 1
               
               If (j <> 1 And (Ji.PRODUCTION_TYPE <> TempProductionType)) Then
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(SumProductionType, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + SumProductionType
                  Total2(k) = Total2(k) + SumProductionType
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff(SumProductionType * 100, ProductAmount))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  SumProductionType = 0
               End If
               
               TempProductionType = Ji.PRODUCTION_TYPE
               TempProductionTypeName = Ji.PRD_TYPE_SHORT_CODE
               
               Set TempJi = GetObject("CJobItem", m_SumInputOutPutLostMain, Trim(m_JobItem.PART_ITEM_ID & "-" & m_JobItem.DOCUMENT_DATE & "-" & Ji.PART_ITEM_ID & "-" & Ji.LOST_ID & "-" & m_JobItem.WEIGHT_AMOUNT & "-" & TempProductionType), True)
               Set TempJiEx = GetObject("CJobItem", m_SumInputOutPutLostNoMain, Trim(m_JobItem.PART_ITEM_ID & "-" & m_JobItem.DOCUMENT_DATE & "-" & Ji.PART_ITEM_ID & "-" & Ji.LOST_ID & "-" & m_JobItem.WEIGHT_AMOUNT & "-" & TempProductionType), True)
               If CheckHaveCollumnDetail(Ji) Then
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(TempJi.TX_AMOUNT + TempJiEx.TX_AMOUNT, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + TempJi.TX_AMOUNT + TempJiEx.TX_AMOUNT
                  Total2(k) = Total2(k) + TempJi.TX_AMOUNT + TempJiEx.TX_AMOUNT
               End If
               
               SumProductionType = SumProductionType + TempJi.TX_AMOUNT + TempJiEx.TX_AMOUNT
               
               If CheckHaveCollumnDetail(Ji) Then
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff((TempJi.TX_AMOUNT + TempJiEx.TX_AMOUNT) * 100, ProductAmount))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
               End If
               
               If (j = m_DistinctOutput.Count) Then
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(SumProductionType, 0)
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  Total1(k) = Total1(k) + SumProductionType
                  Total2(k) = Total2(k) + SumProductionType
                  
                  k = k + 1
                  Set HeadCf = m_Headers1.Fields(k)
                  TempStr = FormatNumberToNull(MyDiff(SumProductionType * 100, ProductAmount))
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
                  SumProductionType = 0
               End If
               
            End If
            
         Next
         If mcolParam("SHOW_SUMMARY") <= 0 Then
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Rs.MoveNext
      Wend
      
      If mcolParam("DAIRY_REPORT") = 0 Then
         If mcolParam("SHOW_SUMMARY") > 0 Then
            Call GenerateFooter(PrevKey2, "", "", Total1, m_Details1, 0)
         Else
            Call GenerateFooter("���", "", "", Total1, m_Details1, 0)
            Vsp.FontBold = True
         End If
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
      End If
      
      Call GenerateFooter("�ط��", "", "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateEndDoc
      
      Vsp.TableBorder = tbNone
      Call PatchTable(Vsp, strFormat, ";", 1000) '10
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      TempStr2 = "................................................|................................................;"
      Call Vsp.AddTable(strFormat, "", TempStr2)
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      TempStr2 = "���Ѵ��|����Ǩ�ͺ;"
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
      ElseIf j > 3 Then
         If mcolParam("DAIRY_REPORT") = 0 Then
            TempStr = FormatNumberToNull(MyDiffEx(Tot(j - 1) * 100, Tot(4)))
         Else
            TempStr = FormatNumberToNull(MyDiffEx(Tot(j - 1) * 100, Tot(5)))
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
   'Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.TAX_ID)
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taLeftMiddle
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   
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
Private Function GetBatchNoSet()
Dim TempStr As String
   GetBatchNoSet = ""
   If Len(mcolParam("BATCH_NO_SET")) > 0 Then
      TempStr = mcolParam("BATCH_NO_SET")
      GetBatchNoSet = "("
      While Len(TempStr) > 0
         If InStr(1, TempStr, ",") > 0 Then
            GetBatchNoSet = GetBatchNoSet & Left(TempStr, InStr(1, TempStr, ",") - 1) & ","
            TempStr = Mid(TempStr, InStr(1, TempStr, ",") + 1, Len(TempStr) - InStr(1, TempStr, ","))
         Else
            GetBatchNoSet = GetBatchNoSet & TempStr & ")"
            TempStr = ""
         End If
      Wend
   End If
End Function
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
Private Function CheckHaveCollumnDetail(Ji As CJobItem) As Boolean
   CheckHaveCollumnDetail = False
   If Ji.PRODUCTION_TYPE = mcolParam("SHOW_DETAIL1") Or _
      Ji.PRODUCTION_TYPE = mcolParam("SHOW_DETAIL2") Or _
      Ji.PRODUCTION_TYPE = mcolParam("SHOW_DETAIL3") Or _
      Ji.PRODUCTION_TYPE = mcolParam("SHOW_DETAIL4") Then
      CheckHaveCollumnDetail = True
   End If
End Function
Private Sub GenerateEndDoc()
   If mcolParam("DAIRY_REPORT") > 0 Then
      Vsp.Paragraph = "*��§ҹ����ѹ"
   End If
   If mcolParam("SHOW_SUMMARY") > 0 Then
      Vsp.Paragraph = "*�ʴ�੾����ػ"
   End If
End Sub


