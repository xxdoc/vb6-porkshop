VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling031"
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
Private Const DETAIL_SIZE = 10

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

Private TotalMoves As Collection
Private TotalMoveCmps As Collection
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
   
   Set TotalMoves = New Collection
   Set TotalMoveCmps = New Collection

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
   
   Set TotalMoves = Nothing
   Set TotalMoveCmps = Nothing
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
   
   Call Cf.SetFieldValue(6, "^", "��¡�����´", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�ʹ����ѹ���" & DateToStringExtEx2(mcolParam("TO_RCP_DATE")), ">")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "�ʹ�������", ">")
   Call m_Headers0.AddField(Cf)
   If mcolParam("SHOW_COST") > 0 Then
      Call Cf.SetFieldValue(4, "^", "�鹷ع", ">")
      Call m_Headers0.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "�鹷ع", ">")
      Call m_Headers0.AddField(Cf)
   End If
   Call Cf.SetFieldValue(4, "^", "����", ">")
   Call m_Headers0.AddField(Cf)
   
   Call Cf.SetFieldValue(6, "^", "��¡���Թ���/����", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ӹǹ", ">", , True)
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ", ">", , True)
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_COST") > 0 Then
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ", ">", , True)
      Call m_Headers1.AddField(Cf)
   Else
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "@", ">")
      Call m_Headers1.AddField(Cf)
      Call Cf.SetFieldValue(0, "^", "�ӹǹ�Թ", ">", , True)
      Call m_Headers1.AddField(Cf)
   End If
   
   Call Cf.SetFieldValue(0, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "@", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ", ">", , True)
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
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call GetSaleAmountAparTypeStockCodeFree(TotalMoves, mcolParam("FROM_BILL_DATE"), mcolParam("TO_BILL_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   Call GetSaleAmountAparTypeStockCodeFree(TotalMoveCmps, mcolParam("TO_RCP_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_STOCK_NO"), mcolParam("TO_STOCK_NO"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")", mcolParam("INCLUDE_FREE"))
   
   FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                       FromDate & " " & ToDate & " �ҡ�١��� " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " �֧�١��� " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & " �ҡ�Թ��� " & EmptyToString(mcolParam("FROM_STOCK_NO"), "N/A") & " �֧�Թ��� " & EmptyToString(mcolParam("TO_STOCK_NO"), "N/A")
                                       
                                       
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
Dim m_BillingDoc As CBillingDoc
Dim TempBillingDoc As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(20) As Double
Dim Total2(20) As Double
Dim Total3(20) As Double
Dim SumItem As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim TotalPrice As Double
Dim CostAmount As Double

Dim SaleReturnColl As Collection
Dim TempSaleReturn As CBillingDoc

Dim k As Long
   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set Rs = New ADODB.Recordset
   Set SaleReturnColl = New Collection
      
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   
   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   I = 0
   Set BodyCf = New CReportField
   
   MasterInd = "79"
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
   m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
   m_BillingDoc.FROM_STOCK_NO = mcolParam("FROM_STOCK_NO")
   m_BillingDoc.TO_STOCK_NO = mcolParam("TO_STOCK_NO")
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.FREE_FLAG = StringToFreeFlag(mcolParam("INCLUDE_FREE"))
   Call m_BillingDoc.QueryData(79, Rs, iCount)
   
   I = 0
   
   While Not Rs.EOF
      MasterInd = "76"
      Set m_BillingDoc = New CBillingDoc
      
      Call m_BillingDoc.PopulateFromRS(79, Rs)
      If m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then       ' Merge Record �ͧ㺢��ʴ��ҡѺ��觢ͧ
         Set TempSaleReturn = GetObject("CBillingDoc", SaleReturnColl, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), False)
         If TempSaleReturn Is Nothing Then
            m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE
            Call SaleReturnColl.add(m_BillingDoc, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO))
         End If
      Else
         Set TempSaleReturn = GetObject("CBillingDoc", SaleReturnColl, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & m_BillingDoc.DOCUMENT_TYPE & "-" & m_BillingDoc.STOCK_NO), False)
         If TempSaleReturn Is Nothing Then
            Call SaleReturnColl.add(m_BillingDoc, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & m_BillingDoc.DOCUMENT_TYPE & "-" & m_BillingDoc.STOCK_NO))
         End If
      End If
      Set m_BillingDoc = Nothing
        
      Rs.MoveNext
   Wend
   
   For Each m_BillingDoc In SaleReturnColl
      I = I + 1
      Call m_Details1.ClearField
      If PrevKey1 <> m_BillingDoc.APAR_TYPE_NAME And I <> 1 Then
         Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         If mcolParam("SHOW_DETAIL") = 1 Then
            Vsp.FontBold = True
         End If
         Call Vsp.AddTable(strFormat, "", TempStr2)
         Vsp.FontBold = False
         
         Call GenerateFooter("      ��� " & PrevKey1, Total2, m_Details1)
          Call m_Details1.GetString(1, TempStr1, TempStr2)
          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
          Vsp.FontBold = True
          Call Vsp.AddTable(strFormat, "", TempStr2)
          Vsp.FontBold = False
          
         Call GenerateHeader("�������١���  :  " & m_BillingDoc.APAR_TYPE_NAME, "������١���  :  " & m_BillingDoc.APAR_GROUP_NAME, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
          Call Vsp.AddTable(strFormat, "", TempStr2)
          
          If mcolParam("SHOW_DETAIL") = 1 Then
             Call GenerateHeader("������  :  " & SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE), "", m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Call Vsp.AddTable(strFormat, "", TempStr2)
             ' ���
          End If
          For j = 1 To UBound(Total1)
             Total1(j) = 0
             Total2(j) = 0
          Next j
          
       ElseIf PrevKey2 <> Trim(Str(m_BillingDoc.DOCUMENT_TYPE)) And I <> 1 Then
          Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
          Call m_Details1.GetString(1, TempStr1, TempStr2)
          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
          If mcolParam("SHOW_DETAIL") = 1 Then
            Vsp.FontBold = True
          End If
          Call Vsp.AddTable(strFormat, "", TempStr2)
          Vsp.FontBold = False
         
         If mcolParam("SHOW_DETAIL") = 1 Then
             Call GenerateHeader("������  :  " & SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE), "", m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Call Vsp.AddTable(strFormat, "", TempStr2)
             ' ���
          End If
          For j = 1 To UBound(Total1)
             Total1(j) = 0
          Next j
       ElseIf I = 1 Then
         Call GenerateHeader("�������١���  :  " & m_BillingDoc.APAR_TYPE_NAME, "������١���  :  " & m_BillingDoc.APAR_GROUP_NAME, m_Details1)
         Call m_Details1.GetString(1, TempStr1, TempStr2)
          strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
          Call Vsp.AddTable(strFormat, "", TempStr2)
          
          If mcolParam("SHOW_DETAIL") = 1 Then
             Call GenerateHeader("������  :  " & SellDoctype2Text(m_BillingDoc.DOCUMENT_TYPE), "", m_Details1)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
             strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
             Call Vsp.AddTable(strFormat, "", TempStr2)
          End If
       End If
       
       PrevKey1 = m_BillingDoc.APAR_TYPE_NAME
       PrevKey2 = Trim(Str(m_BillingDoc.DOCUMENT_TYPE))
       
       Call m_Details1.ClearField
      
       Set HeadCf = m_Headers1.Fields(1)
       TempStr = m_BillingDoc.STOCK_DESC & " /" & m_BillingDoc.STOCK_NO
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         Set TempBillingDoc = GetObject("CBillingDoc", TotalMoveCmps, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
         Set TempSaleReturn = GetObject("CBillingDoc", TotalMoveCmps, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & RECEIPT1_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
       ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then ' ��Ѻ�׹
         Set TempBillingDoc = GetObject("CBillingDoc", TotalMoveCmps, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & RETURN_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
       End If
       Set HeadCf = m_Headers1.Fields(2)
       If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         SumItem = TempBillingDoc.TOTAL_AMOUNT + TempSaleReturn.TOTAL_AMOUNT
       ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
         SumItem = -TempBillingDoc.TOTAL_AMOUNT
       End If
       TempStr = FormatNumberToNull(SumItem)
       Total1(2) = Total1(2) + SumItem
       Total2(2) = Total2(2) + SumItem
       Total3(2) = Total3(2) + SumItem
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
        If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         Amt = Amt + TempSaleReturn.TOTAL_PRICE - TempSaleReturn.DISCOUNT_AMOUNT - TempSaleReturn.EXT_DISCOUNT_AMOUNT
       Else
         Amt = -(TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
       End If
       Set HeadCf = m_Headers1.Fields(3)
       TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
       Total1(3) = Total1(3) + MyDiffEx(Amt, SumItem)
       Total2(3) = Total2(3) + MyDiffEx(Amt, SumItem)
       Total3(3) = Total3(3) + MyDiffEx(Amt, SumItem)
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Set HeadCf = m_Headers1.Fields(4)
       TempStr = FormatNumberToNull(Amt)
       Total1(4) = Total1(4) + Amt
       Total2(4) = Total2(4) + Amt
       Total3(4) = Total3(4) + Amt
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       
       If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & INVOICE_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
         Set TempSaleReturn = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & RECEIPT1_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
       Else
         Set TempBillingDoc = GetObject("CBillingDoc", TotalMoves, Trim(m_BillingDoc.APAR_TYPE_NAME & "-" & RETURN_DOCTYPE & "-" & m_BillingDoc.STOCK_NO), True)
       End If
       Set HeadCf = m_Headers1.Fields(5)
       If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         SumItem = TempBillingDoc.TOTAL_AMOUNT + TempSaleReturn.TOTAL_AMOUNT
       Else
         SumItem = -TempBillingDoc.TOTAL_AMOUNT
       End If
       TempStr = FormatNumberToNull(SumItem)
       Total1(5) = Total1(5) + SumItem
       Total2(5) = Total2(5) + SumItem
       Total3(5) = Total3(5) + SumItem
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         Amt = TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT
         Amt = Amt + TempSaleReturn.TOTAL_PRICE - TempSaleReturn.DISCOUNT_AMOUNT - TempSaleReturn.EXT_DISCOUNT_AMOUNT
       Else
         Amt = -(TempBillingDoc.TOTAL_PRICE - TempBillingDoc.DISCOUNT_AMOUNT - TempBillingDoc.EXT_DISCOUNT_AMOUNT)
       End If
       Set HeadCf = m_Headers1.Fields(6)
       TempStr = FormatNumberToNull(MyDiffEx(Amt, SumItem))
       Total1(6) = Total1(6) + MyDiffEx(Amt, SumItem)
       Total2(6) = Total2(6) + MyDiffEx(Amt, SumItem)
       Total3(6) = Total3(6) + MyDiffEx(Amt, SumItem)
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Set HeadCf = m_Headers1.Fields(7)
       TempStr = FormatNumberToNull(Amt)
       Total1(7) = Total1(7) + Amt
       Total2(7) = Total2(7) + Amt
       Total3(7) = Total3(7) + Amt
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Set HeadCf = m_Headers1.Fields(8)
       TempStr = FormatNumberToNull(SumItem)
       Total1(8) = Total1(8) + SumItem
       Total2(8) = Total2(8) + SumItem
       Total3(8) = Total3(8) + SumItem
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         CostAmount = TempBillingDoc.TOTAL_INCLUDE_PRICE
         CostAmount = CostAmount + (TempSaleReturn.TOTAL_INCLUDE_PRICE)
      Else
         CostAmount = -(m_BillingDoc.TOTAL_INCLUDE_PRICE)
      End If
      Set HeadCf = m_Headers1.Fields(9)
      TempStr = FormatNumberToNull(MyDiffEx(CostAmount, SumItem))
      Total1(9) = Total1(9) + MyDiffEx(CostAmount, SumItem)
      Total2(9) = Total2(9) + MyDiffEx(CostAmount, SumItem)
      Total3(9) = Total3(9) + MyDiffEx(CostAmount, SumItem)
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
       
      Set HeadCf = m_Headers1.Fields(10)
      TempStr = FormatNumberToNull(CostAmount)
      Total1(10) = Total1(10) + CostAmount
      Total2(10) = Total2(10) + CostAmount
      Total3(10) = Total3(10) + CostAmount
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
      Call m_Details1.AddField(BodyCf)
       
       Set HeadCf = m_Headers1.Fields(11)
       TempStr = FormatNumberToNull(SumItem)
       Total1(11) = Total1(11) + SumItem
       Total2(11) = Total2(11) + SumItem
       Total3(11) = Total3(11) + SumItem
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Set HeadCf = m_Headers1.Fields(12)
       TempStr = FormatNumberToNull(MyDiffEx(Amt - CostAmount, SumItem))
       Total1(12) = Total1(12) + MyDiffEx(Amt - CostAmount, SumItem)
       Total2(12) = Total2(12) + MyDiffEx(Amt - CostAmount, SumItem)
       Total3(12) = Total3(12) + MyDiffEx(Amt - CostAmount, SumItem)
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       Set HeadCf = m_Headers1.Fields(13)
       TempStr = FormatNumberToNull(Amt - CostAmount)
       Total1(13) = Total1(13) + Amt - CostAmount
       Total2(13) = Total2(13) + Amt - CostAmount
       Total3(13) = Total3(13) + Amt - CostAmount
       Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
       Call m_Details1.AddField(BodyCf)
       
       If mcolParam("SHOW_DETAIL") = 1 Then
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
      End If
   Next m_BillingDoc
   
   Call GenerateFooter(SellDoctype2Text(Val(PrevKey2)), Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("      ��� " & PrevKey1, Total1, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateFooter("                        �ط��", Total3, m_Details1)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   Call GenerateEndDoc
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set m_BillingDoc = Nothing
   Set TempBillingDoc = Nothing
   Set TotalMoves = Nothing
   Set SaleReturnColl = Nothing
   MasterInd = "1"
   
   genDoc = True
   Exit Function
   
ErrHandler:
   MasterInd = "1"
   mstrErrMsg = "Error(" & RName & ")" & err.Number & " : " & err.Description
   Set Rs = Nothing
End Function
Private Sub GenerateFooter(txt1 As String, Tot() As Double, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempAmount As Double
Dim TempSale As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf HeadCf.SumFlag Then
         TempStr = FormatNumberToNull(Tot(j))
      Else
         TempStr = FormatNumberToNull(MyDiff(Tot(j + 1), Tot(j - 1)))
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
   Call Vsp.AddTable(strFormat, "", "|˹�ҷ��|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExtEx2(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
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
Private Sub GenerateHeader(txt1 As String, txt2 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String
Dim SumWidth As Double
Dim tbdt As TableBorderSettings
   Call Details.ClearField
   Set BodyCf = New CReportField
      
   tbdt = Vsp.TableBorder
   Vsp.TableBorder = tbNone
   SumWidth = 0
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 2 Then
         TempStr = txt1
         SumWidth = SumWidth + HeadCf.Width
         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
         SumWidth = 0
      ElseIf j = m_Headers0.Fields.Count Then
         TempStr = txt2
         SumWidth = SumWidth + HeadCf.Width
         Call BodyCf.SetFieldValue(SumWidth, "<", TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
         SumWidth = 0
      Else
         SumWidth = SumWidth + HeadCf.Width
      End If
   Next HeadCf
      
   Set BodyCf = Nothing
   Vsp.TableBorder = tbdt
End Sub
Private Sub GenerateEndDoc()
   If mcolParam("INCLUDE_FREE") > 0 Then
      Vsp.Paragraph = "*�����¡�âͧ��"
   End If
End Sub

