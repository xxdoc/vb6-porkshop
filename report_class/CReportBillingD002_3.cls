VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBillingD002_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBillingD002"
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

Private CreditBalances As Collection
Private TotalMoves As Collection

Private m_PaidAmounts As Collection
Private m_day1 As Collection
Private m_day2 As Collection
Private day1 As CNumber
Private day2 As CNumber

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
   
   Set CreditBalances = New Collection
   Set TotalMoves = New Collection
   
   Set m_PaidAmounts = New Collection
   Set m_day1 = New Collection
   Set m_day2 = New Collection
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
   Set CreditBalances = Nothing
   Set TotalMoves = Nothing
   
   Set m_PaidAmounts = Nothing
   Set m_day1 = Nothing
   Set m_day2 = Nothing
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
Dim TempDateString1 As String
Dim DateCount1 As Long
Dim DateColum1 As Long
Dim TempDateString2 As String
Dim DateCount2 As Long
Dim DateColum2 As Long
Dim DateLoop As Long


   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   'Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   
   Set Cf = New CReportField
   
   TempDateString1 = Trim(Replace(mcolParam("PERIOD_DATE1"), "(", ""))
   Call CalculateColumnPeriod(TempDateString1, DateColum1)
   TempDateString1 = Trim(Replace(TempDateString1, ")", ""))
   
   TempDateString2 = Trim(Replace(mcolParam("PERIOD_DATE2"), "(", ""))
   Call CalculateColumnPeriod(TempDateString2, DateColum2)
   TempDateString2 = Trim(Replace(TempDateString2, ")", ""))
   
   Call Cf.SetFieldValue(8, "^", "�١���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(DateColum1 * 2, "^", "�Фú��˹�", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(DateColum2 * 2, "^", "�Թ��˹�", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "�ʹ��ҧ���", "<")
   Call m_Headers0.AddField(Cf)
   Call Cf.SetFieldValue(4, "^", "��ǹ����˹��ҧ��ǹ", "<")
   Call m_Headers0.AddField(Cf)
  
   Call Cf.SetFieldValue(2, "^", "�ú��˹�", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "㺡ӡѺ", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ѹ���", "<")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "��ѡ�ҹ���", "^")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(0, "^", "��ҧ���", "^")
   Call m_Headers1.AddField(Cf)
   
   While Len(TempDateString1) > 0
      Call CalculateDatePeriod(TempDateString1, DateCount1)
      If DateLoop = 0 Then
         Call Cf.SetFieldValue(2, "^", "�Թ " & DateCount1 & " �ѹ", ">")
      Else
         Call Cf.SetFieldValue(2, "^", "���� " & DateCount1 & " �ѹ", ">")
      End If
      Call m_Headers1.AddField(Cf)
      DateLoop = DateLoop + 1
   Wend
   DateLoop = 1
   
   Dim DateCount3 As String
   Dim prevDateCount As Long
   Dim collCount As Long
   While Len(TempDateString2) > 0
      Call CalculateDatePeriod(TempDateString2, DateCount2)
      If DateLoop = DateColum2 Then
         Call Cf.SetFieldValue(2, "^", "�Թ " & DateCount2 & " �ѹ", ">")
         Call m_Headers1.AddField(Cf)
      Else
          If prevDateCount = 0 Then
               DateCount3 = "1-" & DateCount2
               prevDateCount = DateCount2 + 1
               Call Cf.SetFieldValue(2, "^", DateCount3 & " �ѹ", ">")
               Call m_Headers1.AddField(Cf)
         Else
                  DateCount3 = prevDateCount & "-" & (prevDateCount + DateCount2 - 1)
                  prevDateCount = prevDateCount + DateCount2
                  Call Cf.SetFieldValue(2, "^", DateCount3 & " �ѹ", ">")
                  Call m_Headers1.AddField(Cf)
         End If
      End If
      DateLoop = DateLoop + 1
   Wend
   
   Call Cf.SetFieldValue(2, "^", "�ʹ˹�����", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�Թ�(�ѹ)", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ������", ">")
   Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ʹ����", ">")
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
Dim ToRcp As String

   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orLandscape
   Vsp.MarginBottom = 700
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 700
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 100
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   If mcolParam("FROM_BILL_DATE") < 1 Then
      FromDate = ""
      ToDate = "� �ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   Else
      FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
      ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   End If
   
   If mcolParam("TO_RCP_DATE") < 1 Then
      ToRcp = "�Դ�ѹ�Ѻ���з�����"
   Else
      ToRcp = "�֧�ѹ����Ѻ���� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_RCP_DATE")), "N/A")
   End If
   
   If mcolParam("TO_PRINT_DATE") < 1 Then
      ToRcp = ToRcp & "  �Դ Due �֧�ѹ���Ѩ�غѹ"
   Else
      ToRcp = ToRcp & "  �Դ Due �֧�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("TO_PRINT_DATE")), "N/A")
   End If
   
   Call GetPaidAmountByDocID(m_PaidAmounts, mcolParam("FROM_BILL_DATE"), mcolParam("TO_RCP_DATE"), mcolParam("FROM_APAR_CODE"), mcolParam("TO_APAR_CODE"), RECEIPT2_DOCTYPE)
   MODULE_DESC = vbCrLf & glbEnterPrise.ENTERPRISE_NAME & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, ",") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " " & ToRcp & " �ҡ�١��� " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " �֧�١��� " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A") & vbCrLf & _
                                        "�ҡ��ѡ�ҹ��� " & EmptyToString(mcolParam("FROM_SALE_CODE"), "N/A") & " �֧��ѡ�ҹ��� " & EmptyToString(mcolParam("TO_SALE_CODE"), "N/A")
                                       
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
Dim L As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim m_BillingDoc As CBillingDoc
Dim TempBd As CBillingDoc
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim Total3(100) As Double
Dim SumRow As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim Rcp As CRcpCnDn_Item
Dim ShowHeader As Boolean
Dim ShowCus As Boolean
Dim showDetail As Boolean
Dim ShowSale As Boolean
Dim M As Long
Dim TempRcp As Double
Dim TempDebit As Double
Dim TempCreDit As Double
Dim ShortCode As String
Dim PRINT_DATE As Date
Dim day1Loop As Long
Dim day2Loop As Long
Dim day1_temp As CNumber
Dim day2_temp As CNumber
Dim checkHaveTat As Long

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
      Total3(j) = 0
   Next j
   
   mblnHeader = False
   Call StartExportFile(Vsp)
   Vsp.StartDoc
   Call printHeader
   mblnHeader = True
   I = 0
   Set BodyCf = New CReportField
   
   Set m_BillingDoc = New CBillingDoc
   
   m_BillingDoc.BILLING_DOC_ID = -1
   m_BillingDoc.ORDER_BY = 1
   m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
   m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
   m_BillingDoc.APAR_IND = 1
   m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
   m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
   m_BillingDoc.FROM_SALE_CODE = mcolParam("FROM_SALE_CODE")
   m_BillingDoc.TO_SALE_CODE = mcolParam("TO_SALE_CODE")
   If mcolParam("SHORT_CODE") = 0 Then
      ShortCode = ""
   Else
      ShortCode = mcolParam("SHORT_CODE")
   End If
   m_BillingDoc.SHORT_CODE = ShortCode
  Call m_BillingDoc.QueryData(164, Rs, iCount)
   
   I = 0
   k = 0
   L = 0
   M = 0
   TempCreDit = 0
   TempDebit = 0
   TempRcp = 0
   
   If mcolParam("TO_PRINT_DATE") <> -1 Then
      PRINT_DATE = mcolParam("TO_PRINT_DATE")
   Else
      PRINT_DATE = Now
   End If
   
   Call dayColl
   
   While Not Rs.EOF
      Call m_Details1.ClearField
      Call m_BillingDoc.PopulateFromRS(164, Rs)
      
      Set Rcp = GetObject("CRcpCnDn_Item", m_PaidAmounts, Trim(Str(m_BillingDoc.BILLING_DOC_ID)))
      If Round(Rcp.PAID_AMOUNT, 2) <> 0 Then
         Call m_PaidAmounts.Remove(Trim(Str(m_BillingDoc.BILLING_DOC_ID)))
      End If
      
      If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT - Rcp.PAID_AMOUNT
         TempRcp = TempRcp + Rcp.PAID_AMOUNT
      ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
         Amt = m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT - Rcp.PAID_AMOUNT
         TempCreDit = TempCreDit + Rcp.PAID_AMOUNT
      ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
         Amt = m_BillingDoc.PAY_AMOUNT - Rcp.PAID_AMOUNT
      End If
      SumRow = Amt
         
      If Round(SumRow, 2) <> 0 Then   'Amt
         If PrevKey1 <> m_BillingDoc.SALE_CODE And M > 0 Then
            
            If Round(Total1(checkHaveTat), 2) <> 0 Then
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
               Else
                  Call GenerateFooter("", PrevKey4 & "(" & PrevKey3 & ")", I & " �", Total1, m_Details1, 0)            '�١���
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Vsp.FontBold = True
               End If
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               ShowHeader = False
               I = 0
            End If

            If Round(Total3(checkHaveTat), 2) <> 0 Then
               Call GenerateFooter("�����ѡ�ҹ���", " " & PrevKey2, k & " �", Total3, m_Details1, 0)            '�������١���
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Vsp.FontBold = True
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
            End If
            
            k = 0
            I = 0

'            Vsp = "��ѡ�ҹ���  :  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_LAST_NAME & " ) "
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
     '          Total2(j) = 0
               Total3(j) = 0
            Next j

         If PrevKey3 <> m_BillingDoc.APAR_CODE And M > 0 Then
            If Round(Total1(checkHaveTat), 2) <> 0 Then
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
               Else
                  Call GenerateFooter("", PrevKey4 & "(" & PrevKey3 & ")", I & " �", Total1, m_Details1, 0)            '�١���
               End If
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Vsp.FontBold = True
               End If
               Call Vsp.AddTable(strFormat, "", TempStr2)
               Vsp.FontBold = False
               ShowHeader = False
               I = 0
            End If
'
            I = 0

            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
         ElseIf M = 0 Then
            ShowHeader = False
            Vsp = "��ѡ�ҹ���  :  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_CODE & " ) "
            Vsp = "�١���  :  " & m_BillingDoc.APAR_LONG_NAME & " " & m_BillingDoc.APAR_LAST_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
         End If
       End If
       
       If PrevKey1 <> m_BillingDoc.SALE_CODE Then
        ShowSale = False
       End If
       If PrevKey3 <> m_BillingDoc.APAR_CODE Then
         ShowCus = False
       End If
                
         If Round(SumRow, 2) <> 0 Then
            If Not (ShowSale) Then
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Vsp = "��ѡ�ҹ���  :  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_CODE & " ) "
                  Vsp = "     �١���  :  " & m_BillingDoc.APAR_LONG_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
               End If
               ShowHeader = True
               ShowSale = True
               ShowCus = True
            ElseIf Not (ShowCus) Then
               If mcolParam("SHOW_DETAIL") = 1 Then
                  Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
                  Call m_Details1.GetString(1, TempStr1, TempStr2)
                  strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                  Vsp.FontBold = True
                  Call Vsp.AddTable(strFormat, "", TempStr2)
                  Vsp.FontBold = False
'                  Vsp = "��ѡ�ҹ���  :  " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_CODE & " ) "
                  Vsp = "     �١���  :  " & m_BillingDoc.APAR_LONG_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
               End If
               ShowCus = True
               showDetail = True
               I = 0
            End If
            
         PrevKey1 = m_BillingDoc.SALE_CODE
         PrevKey2 = m_BillingDoc.SALE_LONG_NAME
         PrevKey3 = m_BillingDoc.APAR_CODE
'         PrevKey4 = m_BillingDoc.APAR_LONG_NAME

         I = I + 1
         k = k + 1
         L = L + 1
         M = M + 1
'         SumRow = 0
         
         Call m_Details1.ClearField
         Set HeadCf = m_Headers1.Fields(1)
         TempStr = "" & DateToStringExtEx2(m_BillingDoc.Due_Date)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If m_BillingDoc.DOCUMENT_NO = "IVV5608383" Then
            Debug.Print
         End If
         
         Set HeadCf = m_Headers1.Fields(2)
         TempStr = m_BillingDoc.DOCUMENT_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(3)
         TempStr = "" & DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(4)
         TempStr = m_BillingDoc.SALE_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = m_BillingDoc.DOC_ID_BILLS_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            SumRow = Amt
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            SumRow = -Amt
         ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Then
            SumRow = -Amt
         ElseIf m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
            SumRow = Amt
         End If
         
'         Dim daydiff As Long

         day1Loop = 1
         day2Loop = 1
'         daydiff = DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date)
'         Set day1 = New CNumber
         For Each day1 In m_day1
            If day1Loop = 1 Then
                  Set HeadCf = m_Headers1.Fields(5 + day1Loop)
                  If DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date) > day1.Day Then
                     TempStr = FormatNumberToNull(SumRow)
                     Total1(5 + day1Loop) = Total1(5 + day1Loop) + SumRow
                     Total2(5 + day1Loop) = Total2(5 + day1Loop) + SumRow
                     Total3(5 + day1Loop) = Total3(5 + day1Loop) + SumRow
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
                  
            ElseIf day1Loop = m_day1.Count Then
                  Set HeadCf = m_Headers1.Fields(5 + day1Loop)
                  If DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date) >= 0 And DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date) <= day1.Day Then
                     TempStr = FormatNumberToNull(SumRow)
                     Total1(5 + day1Loop) = Total1(5 + day1Loop) + SumRow
                     Total2(5 + day1Loop) = Total2(5 + day1Loop) + SumRow
                     Total3(5 + day1Loop) = Total3(5 + day1Loop) + SumRow
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
            Else
                  Set day1_temp = GetObject("CNumber", m_day1, Str(day1Loop + 1), False)
                  Set HeadCf = m_Headers1.Fields(5 + day1Loop)
                  If DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date) > day1_temp.Day And DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date) <= day1.Day Then
                     TempStr = FormatNumberToNull(SumRow)
                     Total1(5 + day1Loop) = Total1(5 + day1Loop) + SumRow
                     Total2(5 + day1Loop) = Total2(5 + day1Loop) + SumRow
                     Total3(5 + day1Loop) = Total3(5 + day1Loop) + SumRow
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
            End If

            day1Loop = day1Loop + 1
         Next day1
         
         day1Loop = 4 + day1Loop
         
        For Each day2 In m_day2
            If day2Loop = m_day2.Count Then
                  Set HeadCf = m_Headers1.Fields(day1Loop + day2Loop)
                  If DateDiff("D", m_BillingDoc.Due_Date, PRINT_DATE) > day2.Day Then
                     TempStr = FormatNumberToNull(SumRow)
                     Total1(day1Loop + day2Loop) = Total1(day1Loop + day2Loop) + SumRow
                     Total2(day1Loop + day2Loop) = Total2(day1Loop + day2Loop) + SumRow
                     Total3(day1Loop + day2Loop) = Total3(day1Loop + day2Loop) + SumRow
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
            Else
'                  Set day2_temp = GetObject("CNumber", m_day2, Str(day2Loop + 1), False)
                  Set HeadCf = m_Headers1.Fields(day1Loop + day2Loop)
                  If DateDiff("D", m_BillingDoc.Due_Date, PRINT_DATE) <= day2.NxDay And DateDiff("D", m_BillingDoc.Due_Date, PRINT_DATE) >= day2.Day Then
                     TempStr = FormatNumberToNull(SumRow)
                     Total1(day1Loop + day2Loop) = Total1(day1Loop + day2Loop) + SumRow
                     Total2(day1Loop + day2Loop) = Total2(day1Loop + day2Loop) + SumRow
                     Total3(day1Loop + day2Loop) = Total3(day1Loop + day2Loop) + SumRow
                  Else
                     TempStr = ""
                  End If
                  Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
                  Call m_Details1.AddField(BodyCf)
            End If

            day2Loop = day2Loop + 1
         Next day2

         day2Loop = day2Loop + day1Loop
         checkHaveTat = day2Loop
         
         Set HeadCf = m_Headers1.Fields(day2Loop)            '13
         TempStr = FormatNumberToNull(SumRow)
         Total1(day2Loop) = Total1(day2Loop) + SumRow
         Total2(day2Loop) = Total2(day2Loop) + SumRow
         Total3(day2Loop) = Total3(day2Loop) + SumRow
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         day2Loop = day2Loop + 1
         Set HeadCf = m_Headers1.Fields(day2Loop)           ' 14
         If DateDiff("D", m_BillingDoc.Due_Date, PRINT_DATE) > 0 Then
            TempStr = DateDiff("D", PRINT_DATE, m_BillingDoc.Due_Date)
         Else
            TempStr = ""
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
          
         day2Loop = day2Loop + 1
         Set HeadCf = m_Headers1.Fields(day2Loop)                                      ' 15
         If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            TempStr = FormatNumberToNull(m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT)
         ElseIf m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            TempStr = FormatNumberToNull(m_BillingDoc.TOTAL_PRICE + m_BillingDoc.VAT_AMOUNT - m_BillingDoc.DISCOUNT_AMOUNT - m_BillingDoc.EXT_DISCOUNT_AMOUNT)
         ElseIf m_BillingDoc.DOCUMENT_TYPE = CN_DOCTYPE Or m_BillingDoc.DOCUMENT_TYPE = DN_DOCTYPE Then
            TempStr = FormatNumberToNull(m_BillingDoc.PAY_AMOUNT)
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         day2Loop = day2Loop + 1
         Set HeadCf = m_Headers1.Fields(day2Loop)                             '16
         TempStr = FormatNumberToNull(Rcp.PAID_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         showDetail = False

            If mcolParam("SHOW_DETAIL") = 1 Then  ''''And showDetail <> True
               Call m_Details1.GetString(1, TempStr1, TempStr2)
               strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
               Call Vsp.AddTable(strFormat, "", TempStr2)
            End If
            
         Else
            I = I - 1
            k = k - 1
            L = L - 1
         End If
      End If
      Rs.MoveNext
   Wend
   
   If Round(Total1(checkHaveTat), 2) <> 0 Then
      If Not (ShowHeader) Then
         Vsp = "     �١���  :  " & m_BillingDoc.APAR_LONG_NAME & " ( " & m_BillingDoc.APAR_CODE & " ) "
      End If

      If mcolParam("SHOW_DETAIL") = 1 Then
         Call GenerateFooter("����١���", " " & PrevKey4, I & " �", Total1, m_Details1, 0)            '�١���
      Else
         Call GenerateFooter("", PrevKey4 & "(" & PrevKey3 & ")", I & " �", Total1, m_Details1, 0)            '�١���
      End If
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      If mcolParam("SHOW_DETAIL") = 1 Then
         Vsp.FontBold = True
      End If
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
       I = 0
   End If
   
      Call GenerateFooter("�����ѡ�ҹ���", " " & PrevKey2, k & " �", Total3, m_Details1, 0)            '�������١���
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
   
   Call GenerateFooter("", "�ط��", L & " �", Total2, m_Details1, 0)
   Call m_Details1.GetString(1, TempStr1, TempStr2)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", TempStr2)
   Vsp.FontBold = False
   
   'debug.print (TempRcp)
   'debug.print (TempCreDit)
   
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set m_BillingDoc = Nothing
   
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
Dim TempWidth As Double

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   If mcolParam("PRINT_TO_EXCEL") = 1 Then
      Set HeadCf = m_Headers1.Fields(1)
      TempStr = txt1
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
      
      Set HeadCf = m_Headers1.Fields(2)
      TempStr = txt2
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Else
      Set HeadCf = m_Headers1.Fields(1)
      TempWidth = TempWidth + HeadCf.Width
      Set HeadCf = m_Headers1.Fields(2)
      TempWidth = TempWidth + HeadCf.Width
      TempStr = txt1 & " " & txt2
      Call BodyCf.SetFieldValue(TempWidth, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)

   End If
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j >= 3 Then
         If j = 3 Then
            TempStr = Txt3
         Else
            TempStr = FormatNumberToNull(Tot(j))
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      End If
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
   
   Vsp.FontSize = iSize
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
Private Sub GenerateHeader(txt1 As String, txt2 As String, Details As CFieldList)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = txt1
      ElseIf j = 2 Then
         TempStr = txt2
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub

Private Sub CalculateDatePeriod(TempStr As String, DateCount As Long)
Dim TempID As Long
   TempID = InStr(1, TempStr, ",")
   If TempID > 0 Then
      DateCount = Val(Left(TempStr, TempID - 1))
      TempStr = Mid(TempStr, TempID + 1)
   Else
      DateCount = Val(TempStr)
      TempStr = ""
   End If
End Sub

Private Sub CalculateColumnPeriod(TempStr_fx As String, DateColum_fx As Long)
Dim TempStr As String
Dim TempID As Long
  TempStr = TempStr_fx
  While (TempStr <> "")
      TempID = InStr(1, TempStr, ",")
      If TempID <> 0 Then
         TempStr = Mid(TempStr, TempID + 1)
         DateColum_fx = DateColum_fx + 1
      Else
         TempStr = ""
         DateColum_fx = DateColum_fx + 1
      End If
   Wend
End Sub

Private Sub dayColl()
Dim DateCount3 As String
Dim prevDateCount As Long
Dim collCount As Long
Dim DateLoop As Long
Dim DateCount1 As Long
Dim DateColum1 As Long
Dim DateCount2 As Long
Dim DateColum2 As Long
Dim TempDateString1 As String
Dim TempDateString2 As String

   TempDateString1 = Trim(Replace(mcolParam("PERIOD_DATE1"), "(", ""))
   Call CalculateColumnPeriod(TempDateString1, DateColum1)
   TempDateString1 = Trim(Replace(TempDateString1, ")", ""))
   
   TempDateString2 = Trim(Replace(mcolParam("PERIOD_DATE2"), "(", ""))
   Call CalculateColumnPeriod(TempDateString2, DateColum2)
   TempDateString2 = Trim(Replace(TempDateString2, ")", ""))

   While Len(TempDateString1) > 0
      Call CalculateDatePeriod(TempDateString1, DateCount1)
      DateLoop = DateLoop + 1
      Set day1 = New CNumber
      day1.Day = DateCount1
      Call m_day1.add(day1, Str(DateLoop))
   Wend
   
   DateLoop = 1

   While Len(TempDateString2) > 0
      Call CalculateDatePeriod(TempDateString2, DateCount2)
      If DateLoop = DateColum2 Then
         collCount = collCount + 1
         Set day2 = New CNumber
         day2.Day = DateCount2
         day2.NxDay = DateCount2
          Call m_day2.add(day2, Str(collCount))
      Else
          If prevDateCount = 0 Then
               prevDateCount = DateCount2 + 1
               collCount = collCount + 1
                     Set day2 = New CNumber
                     day2.Day = 1
                     day2.NxDay = DateCount2
               Call m_day2.add(day2, Str(collCount))
         Else
                  collCount = collCount + 1
                                    Set day2 = New CNumber
                                     day2.Day = prevDateCount
                                    day2.NxDay = prevDateCount + DateCount2 - 1
                                    Call m_day2.add(day2, Str(collCount))
                  prevDateCount = prevDateCount + DateCount2
         End If
      End If
      DateLoop = DateLoop + 1
   Wend

End Sub

