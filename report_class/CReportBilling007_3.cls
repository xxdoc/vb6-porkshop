VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportBilling007_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportBilling007_3"
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

Private Const TITLE_SIZE = 12
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
Private m_YearSeqs As Collection



' note: this API is declared incorrectly in the VB API Viewer.


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
   
   Call Cf.SetFieldValue(2, "^", "�Ţ���", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "�ѹ���", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(8, "^", "�١���", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, "^", "��ҧ�֧", "<")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, ">", "��Ť��", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, ">", "��ǹŴ", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, ">", "VAT.", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(2, ">", "���������", ">")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "�Ѵ", "^")
   Call m_Headers1.AddField(Cf)
   
   Call Cf.SetFieldValue(1, "^", "��", "^")
   Call m_Headers1.AddField(Cf)
   
   If mcolParam("SHOW_DETAIL") = 1 Then
      Call Cf.SetFieldValue(2, "^", "ʶҹФ�ѧ", "<")
      Call m_Headers0.AddField(Cf)
      
      Call Cf.SetFieldValue(14, "^", "��������´", "<")
      Call m_Headers0.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "<", "�ӹǹ", "<")
      Call m_Headers0.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "<", "�Ҥ�/˹���", "<")
      Call m_Headers0.AddField(Cf)
      
      Call Cf.SetFieldValue(2, "<", "��ǹŴ", "<")
      Call m_Headers0.AddField(Cf)
   
      Call Cf.SetFieldValue(2, "<", "�ӹǹ�Թ", "<")
      Call m_Headers0.AddField(Cf)
   End If
   
   Set Cf = Nothing
            
  Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Call m_Headers0.GetString(1, TempStr1, TempStr1_1)
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
   
   ' note: SetTextCharacterExtra takes a spacing in pixels, which
    '       we get by converting from twips.
    'g_CharSpacing = -0 / Vsp.TwipsPerPixelX
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Vsp.ZoomMode = zmPageWidth
   
    FromDate = "�ҡ�ѹ��� " & EmptyToString(DateToStringExtEx2(mcolParam("FROM_BILL_DATE")), "N/A")
   ToDate = "�֧ " & EmptyToString(DateToStringExtEx2(mcolParam("TO_BILL_DATE")), "N/A")
   
   MODULE_DESC = glbEnterPrise.ENTERPRISE_NAME & " " & AddStringFrontEnd(glbEnterPrise.SHORT_NAME, "( ", " )") & vbCrLf & _
                                       mcolParam("REPORT_NAME") & vbCrLf & _
                                        FromDate & " " & ToDate & " �ҡ�١��� " & EmptyToString(mcolParam("FROM_APAR_CODE"), "N/A") & " �֧�١��� " & EmptyToString(mcolParam("TO_APAR_CODE"), "N/A")
                                       
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
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(10) As Double
Dim Total2(10) As Double
Dim Total3(10) As Double
Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim PrevKey3 As String
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim PrevKey8 As String
Dim Ad As CAddress
Dim Ads As Collection

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
      
   Set Ad = New CAddress
   Set Rs = New ADODB.Recordset
   Set Ads = New Collection
   
   If mcolParam("SHOW_DETAIL") = 1 Then
      Call Ad.SetFieldValue("ADDRESS_ID", -1)
      Call LoadAparMasAddress(Ad, Nothing, Ads)
   End If
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total2)
      Total1(j) = 0
      Total2(j) = 0
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
      m_BillingDoc.DOCUMENT_TYPE = RETURN_DOCTYPE
      
      m_BillingDoc.APAR_IND = 1
      m_BillingDoc.FROM_DATE = mcolParam("FROM_BILL_DATE")
      m_BillingDoc.TO_DATE = mcolParam("TO_BILL_DATE")
      m_BillingDoc.FROM_APAR_CODE = mcolParam("FROM_APAR_CODE")
      m_BillingDoc.TO_APAR_CODE = mcolParam("TO_APAR_CODE")
      
       Call m_BillingDoc.QueryData(17, Rs, iCount)
         
      I = 0

      While Not Rs.EOF
         I = I + 1
         Call m_BillingDoc.PopulateFromRS(17, Rs)
         
         If PrevKey1 <> m_BillingDoc.CUSTOMER_BRANCH_CODE And I > 1 Then
            
            If mcolParam("SHOW_DETAIL") = 1 Then
               Set Ad = GetObject("CAddress", Ads, PrevKey7)
               Vsp = "     �����˵�             :  " & PrevKey6
               Vsp = "      ʶҹ����觢ͧ   :  " & PrevKey8 & " " & Ad.PackAddress
            End If
            
            Call GenerateFooter("      ���", PrevKey4, "( " & PrevKey3 & " )", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            Call GenerateFooter("��� ࢵ��â��   : ", PrevKey2, "", Total2, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
               Total2(j) = 0
            Next j
            
            Vsp = ""
            
            Vsp = "ࢵ��â��        : " & m_BillingDoc.CUSTOMER_BRANCH_NAME
            Vsp = "     ��ѡ�ҹ���    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_CODE & " )"
            
         ElseIf PrevKey3 <> m_BillingDoc.SALE_CODE And I <> 1 Then
            If mcolParam("SHOW_DETAIL") = 1 Then
               Set Ad = GetObject("CAddress", Ads, PrevKey7)
               Vsp = "     �����˵�             :  " & PrevKey6
               Vsp = "      ʶҹ����觢ͧ   :  " & PrevKey8 & " " & Ad.PackAddress
            End If
            Call GenerateFooter("      ���", PrevKey4, "( " & PrevKey3 & " )", Total1, m_Details1, 0)
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Vsp.FontBold = True
            Call Vsp.AddTable(strFormat, "", TempStr2)
            Vsp.FontBold = False
            
            Vsp = ""
            
            For j = 1 To UBound(Total1)
               Total1(j) = 0
            Next j
            
            Vsp = ""
            
            Vsp = "     ��ѡ�ҹ���    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_CODE & " )"
            
         ElseIf I = 1 Then
            Vsp = "ࢵ��â��        : " & m_BillingDoc.CUSTOMER_BRANCH_NAME
            Vsp = "     ��ѡ�ҹ���    : " & m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME & " ( " & m_BillingDoc.SALE_CODE & " )"
         End If
         
         PrevKey1 = m_BillingDoc.CUSTOMER_BRANCH_CODE
         PrevKey2 = m_BillingDoc.CUSTOMER_BRANCH_NAME
         PrevKey3 = m_BillingDoc.SALE_CODE
         PrevKey4 = m_BillingDoc.SALE_LONG_NAME & " " & m_BillingDoc.SALE_LAST_NAME
                  
         If PrevKey5 <> m_BillingDoc.DOCUMENT_NO Then
            
            PrevKey5 = m_BillingDoc.DOCUMENT_NO
            PrevKey6 = m_BillingDoc.NOTE
            PrevKey7 = Trim(Str(m_BillingDoc.BRANCH_ADDRESS))
            PrevKey8 = m_BillingDoc.APAR_NAME & " ( " & m_BillingDoc.CUSTOMER_BRANCH_NAME & " )"
            
            Call m_Details1.ClearField
            
            Set HeadCf = m_Headers1.Fields(1)
            TempStr = m_BillingDoc.DOCUMENT_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(2)
            TempStr = DateToStringExtEx2(m_BillingDoc.DOCUMENT_DATE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(3)
            TempStr = m_BillingDoc.SALE_CODE
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1.Fields(4)
            TempStr = m_BillingDoc.SR_REF_DO_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(5)
            TempStr = FormatNumberToNull(m_BillingDoc.TOTAL_PRICE)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(5) = Total1(5) + m_BillingDoc.TOTAL_PRICE
            Total2(5) = Total2(5) + m_BillingDoc.TOTAL_PRICE
            Total3(5) = Total3(5) + m_BillingDoc.TOTAL_PRICE
            
            Set HeadCf = m_Headers1.Fields(6)
            TempStr = FormatNumberToNull(m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(6) = Total1(6) + m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT
            Total2(6) = Total2(6) + m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT
            Total3(6) = Total3(6) + m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT
            
             Set HeadCf = m_Headers1.Fields(7)
            TempStr = FormatNumberToNull(m_BillingDoc.VAT_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(7) = Total1(7) + m_BillingDoc.VAT_AMOUNT
            Total2(7) = Total2(7) + m_BillingDoc.VAT_AMOUNT
            Total3(7) = Total3(7) + m_BillingDoc.VAT_AMOUNT
             
              Set HeadCf = m_Headers1.Fields(8)
            TempStr = FormatNumberToNull(m_BillingDoc.TOTAL_PRICE - (m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT) + m_BillingDoc.VAT_AMOUNT)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(8) = Total1(8) + m_BillingDoc.TOTAL_PRICE - (m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT) + m_BillingDoc.VAT_AMOUNT
            Total2(8) = Total2(8) + m_BillingDoc.TOTAL_PRICE - (m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT) + m_BillingDoc.VAT_AMOUNT
            Total3(8) = Total3(8) + m_BillingDoc.TOTAL_PRICE - (m_BillingDoc.DISCOUNT_AMOUNT + m_BillingDoc.EXT_DISCOUNT_AMOUNT) + m_BillingDoc.VAT_AMOUNT
            
            Set HeadCf = m_Headers1.Fields(9)
            If m_BillingDoc.PAID_AMOUNT > 0 Then
               TempStr = "Y"
            Else
               TempStr = "N"
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers1.Fields(10)
            TempStr = "" '
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         
         If mcolParam("SHOW_DETAIL") = 1 Then
            Call m_Details1.ClearField
            
            Set HeadCf = m_Headers0.Fields(1)
            TempStr = m_BillingDoc.LOCATION_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers0.Fields(2)
            TempStr = m_BillingDoc.STOCK_DESC & " " & m_BillingDoc.STOCK_NO
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(3)
            TempStr = FormatNumberToNull(m_BillingDoc.ITEM_AMOUNT_DOC, 0) & " " & m_BillingDoc.UNIT_CHANGE_NAME
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(4)
            TempStr = FormatNumberToNull(m_BillingDoc.AVG_PRICE_DOC)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(5)
            TempStr = FormatNumberToNull(m_BillingDoc.DISCOUNT_AMOUNT_DOC + m_BillingDoc.EXT_DISCOUNT_AMOUNT_DOC)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Set HeadCf = m_Headers0.Fields(6)
            TempStr = FormatNumberToNull(m_BillingDoc.TOTAL_PRICE_DOC)
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         End If
         Rs.MoveNext
      Wend
      
      Set Ad = GetObject("CAddress", Ads, PrevKey7)
      Vsp = "     �����˵�             :  " & PrevKey6
      Vsp = "      ʶҹ����觢ͧ   :  " & PrevKey8 & " " & Ad.PackAddress
            
       Call GenerateFooter("     ���", PrevKey3, "( " & PrevKey4 & " )", Total1, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("��� ࢵ��â�� : ", PrevKey2, "", Total2, m_Details1, 0)
      Call m_Details1.GetString(1, TempStr1, TempStr2)
      strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
      Vsp.FontBold = True
      Call Vsp.AddTable(strFormat, "", TempStr2)
      Vsp.FontBold = False
      
      Call GenerateFooter("      �ط��", "", "", Total3, m_Details1, 0)
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
   Set Ad = Nothing
   Set Ads = Nothing
   
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
   
   For j = 1 To 4
      Set HeadCf = m_Headers1.Fields(j)
      TempWidth = TempWidth + HeadCf.Width
   Next j
   TempStr = txt1 & " " & txt2 & " " & Txt3
   Call BodyCf.SetFieldValue(TempWidth, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j >= 5 And j <= 8 Then
         TempStr = FormatNumberToNull(Tot(j))
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.BodyAlign)
         Call Details.AddField(BodyCf)
      ElseIf j > 8 Then
         TempStr = ""
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
