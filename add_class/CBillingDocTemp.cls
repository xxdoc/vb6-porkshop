VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillingDocTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public DocItems As Collection
Public RcpCnDnItems As Collection
Public BillingSubTracts As Collection
Public BillingAdditions As Collection

Public Payments As Collection
Public TransferItems As Collection

Public CollBillingDoc As Collection                                     'สำหรับเซ็ต Collection ลูกของตัวเอง ซึ่งเป็น Class BillingDoc

Public CheckTotalAmount As Boolean
Public CheckBillsFlag As Boolean

' BILLING_DOC---------------------------------------------------------------------------------------------------------------------------------
Public DOCUMENT_TYPE As SELL_BILLING_DOCTYPE
Public DOCUMENT_TYPE_RCP As SELL_BILLING_DOCTYPE

Public BILLING_DOC_ID As Long
Public APAR_MAS_ID As Long
Public SALE_BY As Long
Public ENTERPRISE_ADDRESS_ID As Long
Public BILLING_ADDRESS_ID  As Long
Public BRANCH_ADDRESS  As Long
Public DOCUMENT_SUB_TYPE As Long
Public BILLING_DOC_PACK  As Long
Public INVENTORY_DOC_ID As Long
Public SR_REF_DO_ID   As Long

Public DOC_ID_PAID As Double
Public OLD_EXT_DISCOUNT_PERCENT  As Double
Public VAT_PERCENT As Double
Public EXT_DISCOUNT_PERCENT As Double
Public RCP_CASH_TRAN As Double
Public RCP_DIF As Double
Public CASH_PMT As Double
Public CHEQUE_PMT As Double
Public BANKTRF_PMT As Double
Public FEE_AMOUNT As Double
Public TOTAL_PRICE As Double
Public DISCOUNT_AMOUNT As Double
Public EXT_DISCOUNT_AMOUNT As Double
Public TOTAL_AMOUNT As Double
Public VAT_AMOUNT As Double
Public CREDIT As Double
Public PAY_AMOUNT As Double
Public PAID_AMOUNT As Double
Public PAID_DISCOUNT As Double

Public SUBTRACT_AMOUNT As Double
Public ADDITION_AMOUNT As Double
Public DEBIT_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public WH_AMOUNT As Double
Public WH_PERCENT  As Double

Public CUS_PO As String
Public CANCEL_FLAG  As String
Public NOTE As String
Public DOCUMENT_NO As String
Public DO_ADDRESS  As String
Public DOC_ID_BILLS_DUE_DATE As String
Public DOC_ID_BILLS_NO  As String
Public YYYYMM As String
Public DOC_ID_NO As String
Public SR_REF_DO_NO As String

Public DOC_ID_DATE As Date
Public DOCUMENT_DATE As Date
Public Due_Date As Date

' BILLING_DOC---------------------------------------------------------------------------------------------------------------------------------
' DOC_ITEM---------------------------------------------------------------------------------------------------------------------------------
Public DOC_ITEM_ID As Long
Public PO_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID  As Long

Public AVG_PRICE_DOC As Double
Public ITEM_AMOUNT_DOC As Double
Public AVG_PRICE As Double
Public ITEM_AMOUNT  As Double
Public UNIT_MULTIPLE  As Double
Public DISCOUNT_AMOUNT_DOC As Double
Public TOTAL_PRICE_DOC As Double
Public EXT_DISCOUNT_AMOUNT_DOC As Double

Public PO_NO As String
' DOC_ITEM---------------------------------------------------------------------------------------------------------------------------------
' RCPCNDN_ITEM---------------------------------------------------------------------------------------------------------------------------------
Public CNDN_REASON  As Long

Public CNDN_AMOUNT As Double
Public DOC_ID1_PAID  As Double
' RCPCNDN_ITEM---------------------------------------------------------------------------------------------------------------------------------
' CHEQUE---------------------------------------------------------------------------------------------------------------------------------
Public CHEQUE_NO  As String
Public EFFECTIVE_DATE As Date
Public CHEQUE_DATE As Date

' CHEQUE---------------------------------------------------------------------------------------------------------------------------------
' STOCK_CODE---------------------------------------------------------------------------------------------------------------------------------
Public STOCK_TYPE  As Long
Public UNIT_CHANGE_ID As Long

Public STOCK_NO As String
Public STOCK_GROUP_CODE As String
Public STOCK_DESC As String
Public PART_DESC  As String
Public UNIT_AMOUNT As String
' STOCK_CODE---------------------------------------------------------------------------------------------------------------------------------
' APAR_MAS---------------------------------------------------------------------------------------------------------------------------------
Public APAR_IND As Long

Public APAR_CODE As String
Public APAR_NAME As String
Public ADD_BRANCH_NAME  As String
Public APAR_MAS_GROUP_CODE As String
Public APAR_MAS_GROUP_NAME As String
Public TAX_ID As String
Public APAR_TYPE As Long
' APAR_MAS---------------------------------------------------------------------------------------------------------------------------------
' MASTER_REF---------------------------------------------------------------------------------------------------------------------------------
Public UNIT_CHANGE_NAME As String
Public APAR_TYPE_CODE As String
Public UNIT_NAME As String
Public DOCUMENT_SUB_TYPE_NAME  As String
Public UNIT_TRAN_NAME  As String
Public STOCK_TYPE_CODE As String
Public STOCK_TYPE_NAME As String
Public STOCK_GROUP_NAME As String
Public STOCK_TYPE_SUB_CODE As String
Public STOCK_TYPE_SUB_NAME As String
Public APAR_TYPE_NAME As String
Public APAR_GROUP_NAME As String
Public LOCATION_NO As String

Public LOCATION_NAME As String
Public LOCATION_CODE As String
Public SHORT_CODE As String
Public STOCK_GROUP As Long
' MASTER_REF---------------------------------------------------------------------------------------------------------------------------------
' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Public SHOW_ALL_FLAG As String
Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public FROM_APAR_CODE As String
Public TO_APAR_CODE As String
Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String
Public FROM_SALE_CODE  As String
Public TO_SALE_CODE  As String
Public FROM_LOCATION_NO As String
Public TO_LOCATION_NO As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
' LOT_ITEM --------------------------------------------------------------------------------------------------------------------------------------------\
' LOT_ITEM --------------------------------------------------------------------------------------------------------------------------------------------
' EMPLOYEE --------------------------------------------------------------------------------------------------------------------------------------------
Public SALE_CODE As String
Public SALE_NAME As String
' EMPLOYEE --------------------------------------------------------------------------------------------------------------------------------------------
Public CNDN_REASON_NAME As String
Public CONFIG_DOC_TYPE As Long
Public RUNNING_NO  As Long
Public CHECK_BILLS_FLAG  As Boolean
Public DOCUMENT_TYPE_SET  As String
Public NOT_SHOW_RETURN  As String
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   Set DocItems = New Collection
   Set RcpCnDnItems = New Collection
   Set BillingSubTracts = New Collection
   Set BillingAdditions = New Collection
   
   Set Payments = New Collection
   Set TransferItems = New Collection
   
   Set CollBillingDoc = New Collection
   CANCEL_FLAG = "N"                                              'ตั้งแบบนี้หมายถึงให้ DEFAULT เป็น N นั้นเอง
   
End Sub
Private Sub Class_Terminate()
   Set DocItems = Nothing
   Set RcpCnDnItems = Nothing
   Set BillingSubTracts = Nothing
   Set BillingAdditions = Nothing
   Set Payments = Nothing
   Set TransferItems = Nothing
   
   Set CollBillingDoc = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
Dim iCount As Long
Dim TempRs As ADODB.Recordset
         
   Set TempRs = New ADODB.Recordset
   Dim Ct As CCashTran '
   Dim Cq As CCheque
   Set Ct = New CCashTran
   Call Ct.SetFieldValue("BILLING_DOC_ID", BILLING_DOC_ID)
   Call Ct.QueryData(1, TempRs, iCount)
   Set Ct = Nothing
   While Not TempRs.EOF
      Set Ct = New CCashTran
      Call Ct.PopulateFromRS(1, TempRs)
      Call Ct.DeleteDataBillingDoc
      If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
         Set Cq = New CCheque
         Call Cq.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
         Call Cq.DeleteData
         Set Cq = Nothing
      End If
      Set Ct = Nothing
      Call TempRs.MoveNext
   Wend
   Set Cq = Nothing
   Set Ct = Nothing
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   
   SQL = "DELETE FROM DOC_ITEM WHERE  BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM RCPCNDN_ITEM WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_SUBTRACT WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_ADDITION WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM LOT_ITEM WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM INVENTORY_DOC WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
End Sub
Public Sub DeleteDataForCNDN()
Dim SQL As String
Dim iCount As Long
   
   SQL = "DELETE FROM RCPCNDN_ITEM WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub DeleteDataForSR()
Dim SQL As String
Dim iCount As Long
   
   SQL = "DELETE FROM PRINT_LABEL PLA WHERE PLA.DOC_ITEM_ID IN (SELECT DI.DOC_ITEM_ID FROM DOC_ITEM DI  WHERE   DI.BILLING_DOC_ID  = " & BILLING_DOC_ID & ")  "
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM DOC_ITEM_LINK DIK WHERE DIK.DOC_ITEM_ID IN (SELECT DI.DOC_ITEM_ID FROM DOC_ITEM DI  WHERE   DI.BILLING_DOC_ID  = " & BILLING_DOC_ID & ")  "
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM DOC_ITEM WHERE  BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub DeleteDataForIV()
Dim SQL As String
Dim iCount As Long
   
   SQL = "DELETE FROM PRINT_LABEL PLA WHERE PLA.DOC_ITEM_ID IN (SELECT DI.DOC_ITEM_ID FROM DOC_ITEM DI  WHERE   DI.BILLING_DOC_ID  = " & BILLING_DOC_ID & ")  "
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM DOC_ITEM_LINK DIK WHERE DIK.DOC_ITEM_ID IN (SELECT DI.DOC_ITEM_ID FROM DOC_ITEM DI  WHERE   DI.BILLING_DOC_ID  = " & BILLING_DOC_ID & ")  "
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM DOC_ITEM WHERE  BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID  = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM LOT_ITEM_LINK LK WHERE LK.EXPORT_LOT_ITEM_ID IN (SELECT LT.LOT_ITEM_ID FROM LOT_ITEM LT  WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID & ")"
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM LOT_ITEM WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM INVENTORY_DOC WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim NewStr As String
Dim SubLen As Long
Dim TempDate As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, APM.APAR_CODE,APM.APAR_NAME, APM.TAX_ID "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.EMP_NAME SALE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME DOCUMENT_SUB_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.DOCUMENT_SUB_TYPE = MR3.KEY_ID) "
      
      MustCount = True
      If CheckBillsFlag Then
         WhereStr = "WHERE (      ((SELECT COUNT(*) FROM RCPCNDN_ITEM RCP WHERE RCP.BILLING_DOC_ID = NM.BILLING_DOC_ID) > (SELECT COUNT(*) FROM RCPCNDN_ITEM RCP WHERE RCP.BILLS_ID = NM.BILLING_DOC_ID)) "
         WhereStr = WhereStr & " AND    (  (SELECT COUNT(*) FROM RCPCNDN_ITEM RCP1 LEFT OUTER JOIN RCPCNDN_ITEM RCP2 ON (RCP1.DOC_ID_BILLS = RCP2.DOC_ID_RCP) WHERE RCP1.BILLING_DOC_ID = NM.BILLING_DOC_ID AND RCP2.BILLING_DOC_ID IS NULL)  >0)  ) "
      End If
      
   ElseIf Ind = 3 Then 'สำหับ report billing002
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT , NM.DOCUMENT_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_CHANGE_ID = MR2.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (APM.APAR_IND =  1 AND SC.STOCK_NO IS NOT NULL ) " & vbCrLf
    ElseIf Ind = 5 Then 'สำหับ reportCommission001
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT,SUM(DI.TOTAL_PRICE) TOTAL_PRICE,NM.SALE_BY "
      SelectStr1 = SelectStr1 & ",SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.GROUP_COM_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.SALE_BY >  0) " & vbCrLf
   ElseIf Ind = 9 Then 'reportbillingdoc008
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT , NM.DOCUMENT_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_TRAN_NAME  "
      SelectStr1 = SelectStr1 & ",DI.UNIT_MULTIPLE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (DI.UNIT_TRAN_ID = MR2.KEY_ID) "
   ElseIf Ind = 10 Then 'reportbillingdoc009
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT , NM.DUE_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_TRAN_NAME  "
      SelectStr1 = SelectStr1 & ",DI.UNIT_MULTIPLE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (DI.UNIT_TRAN_ID = MR2.KEY_ID) "
      TempStr = ""
   ElseIf Ind = 11 Then 'reportbillingdocD001
      SelectStr1 = "SELECT DISTINCT MR2.KEY_CODE APAR_TYPE_CODE,MR2.KEY_NAME APAR_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      TempStr = ""
   ElseIf Ind = 12 Then 'LoadSaleAmount      'GetCreditBalance
      SelectStr1 = "SELECT SUM(NM.TOTAL_PRICE) TOTAL_PRICE,SUM(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.PAID_AMOUNT) PAID_AMOUNT "  'รับชำระ เพิ่มหนี้ ลดหนี้
      SelectStr1 = SelectStr1 & ",SUM(NM.DEBIT_AMOUNT) DEBIT_AMOUNT "  'เพิ่มหนี้ จากการรับชำระ
      SelectStr1 = SelectStr1 & ",SUM(NM.CREDIT_AMOUNT) CREDIT_AMOUNT "  'ลดหนี้ จากการรับชำระ
      
      SelectStr1 = SelectStr1 & ", NM.APAR_MAS_ID,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      TempStr = ""
   ElseIf Ind = 13 Then 'reportbillingdocD002 and reportbillingdocD003,frmProcessEndYear
      SelectStr1 = "SELECT NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE,NM.BILLING_DOC_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE APAR_TYPE_CODE,MR2.KEY_NAME APAR_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.DUE_DATE,EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      TempStr = ""
      
   ElseIf Ind = 14 Then 'reportbillingdocD004
      SelectStr1 = "SELECT NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ",SUM(RCP.PAID_AMOUNT) PAID_AMOUNT "
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE,APM.APAR_TYPE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP   ) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      TempStr = ""
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM1.DOCUMENT_TYPE = " & RECEIPT2_DOCTYPE & " OR NM1.DOCUMENT_TYPE IS NULL ) "
   ElseIf Ind = 15 Then 'Reportbillingdoc007_1
      SelectStr1 = "SELECT NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.VAT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT  "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.SR_REF_DO_NO "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.NOTE "
      
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",DI.ITEM_AMOUNT ITEM_AMOUNT_DOC,DI.AVG_PRICE AVG_PRICE_DOC,DI.DISCOUNT_AMOUNT DISCOUNT_AMOUNT_DOC,DI.TOTAL_PRICE TOTAL_PRICE_DOC "
      SelectStr1 = SelectStr1 & ",DI.EXTRA_DISCOUNT_AMOUNT  EXT_DISCOUNT_AMOUNT_DOC "
      SelectStr1 = SelectStr1 & ",LC.KEY_CODE LOCATION_NO "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (DI.LOCATION_ID = LC.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.UNIT_CHANGE_ID = MR5.KEY_ID) "
      
      TempStr = ""
   ElseIf Ind = 17 Then 'Reportbillingdoc007_3
      SelectStr1 = "SELECT NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.VAT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT  "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",NM.SR_REF_DO_NO "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.NOTE,NM.BRANCH_ADDRESS "
      
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",DI.ITEM_AMOUNT ITEM_AMOUNT_DOC,DI.AVG_PRICE AVG_PRICE_DOC,DI.DISCOUNT_AMOUNT DISCOUNT_AMOUNT_DOC,DI.TOTAL_PRICE TOTAL_PRICE_DOC "
      SelectStr1 = SelectStr1 & ",DI.EXTRA_DISCOUNT_AMOUNT  EXT_DISCOUNT_AMOUNT_DOC "
      SelectStr1 = SelectStr1 & ",LC.KEY_CODE LOCATION_NO "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (DI.LOCATION_ID = LC.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.UNIT_CHANGE_ID = MR5.KEY_ID) "
      
      TempStr = ""
      
   ElseIf Ind = 18 Or Ind = 19 Or Ind = 129 Then  'Reportbillingdoc010 and Reportbillingdoc011
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE,NM.VAT_PERCENT,NM.EXT_DISCOUNT_PERCENT,NM.NOTE,NM.REFER_TEXT "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",DI.ITEM_AMOUNT,DI.UNIT_MULTIPLE,DI.AVG_PRICE ,DI.DISCOUNT_AMOUNT,DI.EXTRA_DISCOUNT_AMOUNT EXT_DISCOUNT_AMOUNT,DI.PO_NO "
      SelectStr1 = SelectStr1 & ",DI.TOTAL_PRICE "
      
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME UNIT_TRAN_NAME "
      SelectStr1 = SelectStr1 & ",MR6.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (DI.UNIT_TRAN_ID = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (SC.UNIT_CHANGE_ID = MR6.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 20 Then 'reportbillingdoc012
      SelectStr1 = "SELECT MAX(NM.TOTAL_PRICE) TOTAL_PRICE,MAX(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.PAY_AMOUNT) PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",RCP.PAID_DISCOUNT PAID_DISCOUNT "                     'ส่วนลดรับชำระ
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE,NM3.DOCUMENT_NO DOC_ID_BILLS_NO,NM3.DUE_DATE DOC_ID_BILLS_DUE_DATE,APM.APAR_TYPE "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP   ) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP1 ON (NM.BILLING_DOC_ID = RCP1.DOC_ID_BILLS   ) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM3 ON (RCP1.BILLING_DOC_ID = NM3.BILLING_DOC_ID ) "
      TempStr = ""
   ElseIf Ind = 21 Then 'GetSaleAmountMonth
      SelectStr1 = "SELECT SUM(NM.TOTAL_PRICE) TOTAL_PRICE,SUM(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.PAID_AMOUNT) PAID_AMOUNT "  'รับชำระ เพิ่มหนี้ ลดหนี้
      SelectStr1 = SelectStr1 & ",SUM(NM.DEBIT_AMOUNT) DEBIT_AMOUNT "  'เพิ่มหนี้ จากการรับชำระ
      SelectStr1 = SelectStr1 & ",SUM(NM.CREDIT_AMOUNT) CREDIT_AMOUNT "  'ลดหนี้ จากการรับชำระ
      
      SelectStr1 = SelectStr1 & ", NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM.YYYYMM "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      TempStr = ""
   ElseIf Ind = 22 Then 'สำหับ reportbillingdoc015
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT,SUM(DI.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",TOP(NM.EXT_DISCOUNT_PERCENT) EXT_DISCOUNT_PERCENT "
      SelectStr1 = SelectStr1 & ",DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",MR6.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME APAR_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",NM.YYYYMM,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (SC.UNIT_CHANGE_ID = MR6.KEY_ID) "
      
   ElseIf Ind = 23 Then ' FrmAddReceiptItem Sub QueryData
      SelectStr1 = "SELECT MAX(NM.TOTAL_PRICE) TOTAL_PRICE,MAX(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.PAY_AMOUNT) PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      If SHOW_ALL_FLAG <> "Y" Then
         SelectStr1 = SelectStr1 & ",SUM(RCP.PAID_AMOUNT) PAID_AMOUNT "
         SelectStr1 = SelectStr1 & ",SUM(RCP.PAID_DISCOUNT) PAID_DISCOUNT "                     'ส่วนลดรับชำระ
      Else
         SelectStr1 = SelectStr1 & ",0 PAID_AMOUNT"
         SelectStr1 = SelectStr1 & ",0 PAID_DISCOUNT"                     'ส่วนลดรับชำระ
      End If
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.BILLING_DOC_ID,NM.REFER_TEXT,NM.CUS_PO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      If SHOW_ALL_FLAG <> "Y" Then
         If DOCUMENT_TYPE_RCP = RECEIPT2_DOCTYPE Then
            FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP   ) "
         Else
            FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID) " 'ใบวางบิล
         End If
         FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
         
         
         TempStr = ""
         
         TempStr = "WHERE "
         If DOCUMENT_TYPE_RCP = RECEIPT2_DOCTYPE Then
            WhereStr = WhereStr & TempStr & " (NM1.DOCUMENT_TYPE = " & DOCUMENT_TYPE_RCP & " OR NM1.DOCUMENT_TYPE IS NULL ) "
         Else 'ใบวางบิลจะคิดทุกอันเลย
            WhereStr = WhereStr & TempStr & " (NM1.DOCUMENT_TYPE IN (" & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & "," & BILLS_DOCTYPE & "," & RECEIPT2_DOCTYPE & ") OR NM1.DOCUMENT_TYPE IS NULL ) "
         End If
      End If
   ElseIf Ind = 24 Then 'สำหับ GetSaleAmountMonthByStockCode
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.YYYYMM,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      
   ElseIf Ind = 25 Then 'สำหับ report billing017
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      'SelectStr1 = SelectStr1 & ",SUM(LI.AVG_PRICE ) AVG_PRICE  "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_NAME,MR3.KEY_NAME UNIT_CHANGE_NAME,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      
'      'FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM_LINK DIK ON (DIK.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
'      'FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (DIK.MAIN_IMPORT_LOT_ITEM_ID = LI.LOT_ITEM_ID) "
      
   ElseIf Ind = 26 Then 'สำหับ report billing018
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      
      TempStr = ""
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO IS NOT NULL ) "
   ElseIf Ind = 27 Then 'สำหับ report billing018
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.YYYYMM,NM.DOCUMENT_TYPE,APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
   ElseIf Ind = 28 Then 'reportbillingD005
      SelectStr1 = "SELECT NM.DOCUMENT_DATE,NM.DOCUMENT_NO,NM.DOCUMENT_TYPE,NM.APAR_MAS_ID "
      SelectStr1 = SelectStr1 & ",NM.TOTAL_PRICE,NM.VAT_AMOUNT,NM.DISCOUNT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.PAID_AMOUNT,NM.DEBIT_AMOUNT,NM.CREDIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.CASH_PMT,NM.CHEQUE_PMT,NM.BANKTRF_PMT "
      SelectStr1 = SelectStr1 & ",NM1.DOCUMENT_DATE DOC_ID_DATE,NM1.DOCUMENT_NO DOC_ID_NO,NM1.DOCUMENT_TYPE DOCUMENT_TYPE_RCP "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT DOC_ID_PAID,RCP.CNDN_AMOUNT "
      SelectStr1 = SelectStr1 & ",CQ.CHEQUE_NO,CQ.CHEQUE_DATE,CQ.EFFECTIVE_DATE "
      SelectStr1 = SelectStr1 & ",RCP1.PAID_AMOUNT DOC_ID1_PAID "
      'SelectStr1 = SelectStr1 & ",0 DOC_ID1_PAID "
      
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN CASH_TRAN CT ON (NM.BILLING_DOC_ID = CT.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE CQ ON (CQ.CHEQUE_ID = CT.CHEQUE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.DOC_ID_RCP = NM1.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP1 ON (NM.BILLING_DOC_ID = RCP1.DOC_ID_RCP) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM3 ON (RCP1.BILLING_DOC_ID  = NM3.BILLING_DOC_ID ) "
      
      TempStr = ""
      
   ElseIf Ind = 29 Then 'สำหับ report billing019
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_NAME,MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,DI.FREE_FLAG "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
   ElseIf Ind = 30 Then 'Reportbillingdoc020
      SelectStr1 = "SELECT NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.VAT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT  "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP) "
      
      TempStr = ""
   ElseIf Ind = 31 Then 'Reportbillingdoc021
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",DI.ITEM_AMOUNT,DI.AVG_PRICE,DI.DISCOUNT_AMOUNT,DI.TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",DI.EXTRA_DISCOUNT_AMOUNT  EXT_DISCOUNT_AMOUNT,DI.UNIT_MULTIPLE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME UNIT_TRAN_NAME "
      SelectStr1 = SelectStr1 & ",DI.PO_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (DI.UNIT_TRAN_ID = MR5.KEY_ID) "
      
      TempStr = ""
   ElseIf Ind = 32 Then 'Reportbillingdoc022
      SelectStr1 = "SELECT NM.DOCUMENT_NO,NM.DOCUMENT_DATE  "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",DI.ITEM_AMOUNT,DI.AVG_PRICE,DI.DISCOUNT_AMOUNT,DI.TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",DI.EXTRA_DISCOUNT_AMOUNT EXT_DISCOUNT_AMOUNT,DI.UNIT_MULTIPLE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME UNIT_TRAN_NAME "
      
     SelectStr1 = SelectStr1 & ",MR1.KEY_NAME LOCATION_NAME ,MR1.KEY_CODE LOCATION_CODE "
       
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (DI.UNIT_TRAN_ID = MR5.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (DI.LOCATION_ID = MR1.KEY_ID) "
      
      TempStr = ""
   ElseIf Ind = 33 Then 'สำหรับ ClearDataBillingDocStockCash
      SelectStr1 = "SELECT DISTINCT DI.DOC_ITEM_ID "
      
      SelectStr2 = "SELECT COUNT(DISTINCT DI.DOC_ITEM_ID) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      TempStr = ""
   ElseIf Ind = 34 Then 'สำหับ report billing019
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",MR4.KEY_CODE STOCK_TYPE_CODE ,MR4.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.STOCK_TYPE = MR4.KEY_ID) "
      
      TempStr = ""
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO IS NOT NULL ) "
      If STOCK_GROUP > 0 Then
         If Len(WhereStr) > 0 Then
            TempStr = "AND "
         Else
            TempStr = "WHERE "
         End If
         WhereStr = WhereStr & TempStr & " (MR4.PARENT_ID = " & STOCK_GROUP & ") " & vbCrLf
      End If
   ElseIf Ind = 35 Then 'สำหับ report billing023
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE ,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.STOCK_TYPE = MR4.KEY_ID) "
      
      TempStr = ""
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO IS NOT NULL ) "
   ElseIf Ind = 36 Then 'GetSaleAmountAparStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_NAME,MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,DI.FREE_FLAG,APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
   ElseIf Ind = 37 Then 'GetSaleAmountBranchEmpAparStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 38 Then 'GetSaleAmountBranchEmpStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' ) " & vbCrLf
   ElseIf Ind = 39 Then 'GetSaleAmountStockCodeDocTypeFreeEx
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' ) " & vbCrLf
      
   ElseIf Ind = 40 Then 'GetDistinctBranchEmpAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
      

   ElseIf Ind = 47 Then 'GetDistinctStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE STOCK_TYPE_CODE ,MR2.KEY_NAME STOCK_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' ) " & vbCrLf
   ElseIf Ind = 48 Then 'GetDistinctBranchAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   
   ElseIf Ind = 49 Then 'GetDistinctBranchAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 50 Then 'GetSaleAmountBranchAparStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 51 Then 'GetSaleAmountBranchAparStockCodeDocTypeFreeEx
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 52 Then 'GetDistinctBranchStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 53 Then 'GetDistinctBranchStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE CUSTOMER_BRANCH_CODE,MR2.KEY_NAME CUSTOMER_BRANCH_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 54 Then 'GetSaleAmountBranchStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 55 Then 'GetSaleAmountBranchStockCodeDocTypeFreeEx
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 56 Then 'GetDistinctEmpAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.JOINT_CODE,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME CUSTOMER_BRANCH_NAME,NM.CUSTOMER_BRANCH,MR5.KEY_CODE CUSTOMER_BRANCH_CODE,EM.JOINT_CODE SALE_JOINT_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = NM.CUSTOMER_BRANCH) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
      
   ElseIf Ind = 57 Then 'GetDistinctEmpAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.JOINT_CODE,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CUSTOMER_BRANCH_NAME,PL.BRANCH_ID CUSTOMER_BRANCH,MR2.KEY_CODE CUSTOMER_BRANCH_CODE,EM.JOINT_CODE SALE_JOINT_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
      
   ElseIf Ind = 58 Then 'GetSaleAmountEmpAparStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 59 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeEx
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 60 Then 'GetDistinctEmpStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.JOINT_CODE,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE SALE_JOINT_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 61 Then 'GetDistinctEmpStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.JOINT_CODE,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE SALE_JOINT_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 62 Then 'GetSaleAmountEmpStockCodeDocTypeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 63 Then 'GetSaleAmountEmpStockCodeDocTypeFreeEx
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 64 Then 'GetDistinctAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 65 Then 'GetDistinctAparStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 66 Then 'GetSaleAmountAparStockCodeDocTypeFreeEx
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 67 Then 'GetSaleAmountAparStockCodeDocTypeFreeEx1
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 69 Then 'Reportbillingdoc025
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT,NM.PAID_AMOUNT,NM.SUBTRACT_AMOUNT,NM.ADDITION_AMOUNT,NM.DEBIT_AMOUNT,NM.CREDIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT,NM.WH_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.RCP_DIF,NM.CASH_PMT,NM.CHEQUE_PMT,NM.BANKTRF_PMT,NM.FEE_AMOUNT "
      
      SelectStr1 = SelectStr1 & ",NM1.DOCUMENT_DATE DOC_ID_DATE,NM1.DOCUMENT_NO DOC_ID_NO,NM1.DOCUMENT_TYPE DOCUMENT_TYPE_RCP "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT DOC_ID_PAID,RCP.CNDN_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.DOC_ID_RCP = NM1.BILLING_DOC_ID ) "
      
      TempStr = ""
   ElseIf Ind = 71 Then 'Reportbillingdoc027
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT,NM.PAID_AMOUNT,NM.SUBTRACT_AMOUNT,NM.ADDITION_AMOUNT,NM.DEBIT_AMOUNT,NM.CREDIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT,NM.WH_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.RCP_DIF,NM.CASH_PMT,NM.CHEQUE_PMT,NM.BANKTRF_PMT,NM.FEE_AMOUNT "
      
      SelectStr1 = SelectStr1 & ",NM1.DOCUMENT_DATE DOC_ID_DATE,NM1.DOCUMENT_NO DOC_ID_NO,NM1.DOCUMENT_TYPE DOCUMENT_TYPE_RCP "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT DOC_ID_PAID,RCP.CNDN_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.DOC_ID_RCP = NM1.BILLING_DOC_ID ) "
      
      TempStr = ""
   ElseIf Ind = 73 Then 'Reportbillingdoc029
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT,NM.PAID_AMOUNT,NM.SUBTRACT_AMOUNT,NM.ADDITION_AMOUNT,NM.DEBIT_AMOUNT,NM.CREDIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT,NM.WH_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.RCP_DIF,NM.CASH_PMT,NM.CHEQUE_PMT,NM.BANKTRF_PMT,NM.FEE_AMOUNT "
      
      SelectStr1 = SelectStr1 & ",NM1.DOCUMENT_DATE DOC_ID_DATE,NM1.DOCUMENT_NO DOC_ID_NO,NM1.DOCUMENT_TYPE DOCUMENT_TYPE_RCP "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT DOC_ID_PAID,RCP.CNDN_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.DOC_ID_RCP = NM1.BILLING_DOC_ID ) "
      
      TempStr = ""
   ElseIf Ind = 74 Then  'frmAddReturnItem      ,   'frmAddPoItem
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_DATE,NM.DOCUMENT_NO,NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",NM.CUS_PO,NM.VAT_PERCENT,NM.EXT_DISCOUNT_PERCENT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      If CheckTotalAmount Then               'frmAddPoItem
         WhereStr = "WHERE ((NM.TOTAL_AMOUNT > (SELECT SUM(DI.ITEM_AMOUNT) FROM DOC_ITEM DI WHERE DI.PO_ID = NM.BILLING_DOC_ID)) "
         WhereStr = WhereStr & " OR (SELECT COUNT(*) FROM DOC_ITEM DI WHERE DI.PO_ID = NM.BILLING_DOC_ID) = 0 ) "
      End If
         
   ElseIf Ind = 75 Then 'GetSaleAmountAparGroupDocTypeStockCodeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (APM.APAR_TYPE = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR4.PARENT_ID = MR5.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 76 Then 'สำหับ report billing030
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
   ElseIf Ind = 77 Then 'GetSaleAmountStockCodeDocType
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.JOINT_CODE,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      WhereStr = ""
   
   ElseIf Ind = 78 Then 'GetSaleAmountAparTypeStockCodeFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (APM.APAR_TYPE = MR4.KEY_ID) "
         
      WhereStr = ""
   ElseIf Ind = 79 Then 'สำหับ report billing031
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME APAR_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
   ElseIf Ind = 80 Then 'สำหรับ GetDistinctStockCode
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT,SC.JOINT_CODE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME STOCK_GROUP_NAME,MR5.KEY_CODE STOCK_GROUP_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = MR2.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
   ElseIf Ind = 83 Then 'สำหับ report billing034 and 'สำหับ report billing035
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME APAR_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
   ElseIf Ind = 84 Then 'GetSaleAmountAparTypeStockCodeDateFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_TYPE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (APM.APAR_TYPE = MR4.KEY_ID) "
      
      WhereStr = " WHERE DI.FREE_FLAG = 'N' "
   ElseIf Ind = 85 Then 'GetSaleAmountStockCodeDocTypeDate
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      WhereStr = ""
   ElseIf Ind = 86 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.YYYYMM,NM.CUSTOMER_BRANCH "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
      'WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 87 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.YYYYMM,PL.BRANCH_ID CUSTOMER_BRANCH  "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 88 Then 'Patch_2007_11_30_1_jill
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & "  ((select sum(bd.total_price) from billing_doc bd where bd.billing_doc_id =  nm.billing_doc_id) <>"
      WhereStr = WhereStr & " (select sum(di.total_price) from doc_item di where di.billing_doc_id =  nm.billing_doc_id) " & vbCrLf
   ElseIf Ind = 89 Then 'Patch_2007_11_30_2_jill
      SelectStr1 = "SELECT NM.BILLING_DOC_ID,NM.DOCUMENT_NO,NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & "  ((select sum(bd.ext_discount_amount) from billing_doc bd where bd.billing_doc_id =  nm.billing_doc_id) <>"
      WhereStr = WhereStr & " (select sum(di.extra_discount_amount) from doc_item di where di.billing_doc_id =  nm.billing_doc_id) " & vbCrLf
   ElseIf Ind = 90 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_PRICE) TOTAL_PRICE,SUM(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.WH_AMOUNT) WH_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(NM.PAY_AMOUNT) PAY_AMOUNT "                'ส่วนเพิ่มหนี้ลดหนี้
      SelectStr1 = SelectStr1 & ",SUM(NM.DEBIT_AMOUNT) DEBIT_AMOUNT "                'ลดหนี้
      SelectStr1 = SelectStr1 & ",SUM(NM.CREDIT_AMOUNT) CREDIT_AMOUNT "                'เพิ่มหนี้
      SelectStr1 = SelectStr1 & ",SUM(NM.SUBTRACT_AMOUNT) SUBTRACT_AMOUNT "                'ส่วนหัก
      SelectStr1 = SelectStr1 & ",SUM(NM.ADDITION_AMOUNT) ADDITION_AMOUNT "                'ส่วนเพิ่ม
      SelectStr1 = SelectStr1 & ", NM.APAR_MAS_ID,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      TempStr = ""
   ElseIf Ind = 91 Then 'GetSaleAmountAparGroupStockCodeDateFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME APAR_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (APM.APAR_TYPE = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR4.PARENT_ID = MR5.KEY_ID) "
               
      'FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM_LINK DIK ON (DIK.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      'FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (DIK.MAIN_IMPORT_LOT_ITEM_ID = LI.LOT_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 92 Then 'สำหรับ report billing035_1
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
   ElseIf Ind = 93 Then 'Reportbillingdoc007_4
      SelectStr1 = "SELECT NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.VAT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT  "
      SelectStr1 = SelectStr1 & ",APM.APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",NM.SR_REF_DO_NO "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.NOTE "
      
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",DI.ITEM_AMOUNT ITEM_AMOUNT_DOC,DI.AVG_PRICE AVG_PRICE_DOC,DI.DISCOUNT_AMOUNT DISCOUNT_AMOUNT_DOC,DI.TOTAL_PRICE TOTAL_PRICE_DOC "
      SelectStr1 = SelectStr1 & ",DI.EXTRA_DISCOUNT_AMOUNT  EXT_DISCOUNT_AMOUNT_DOC "
      SelectStr1 = SelectStr1 & ",LC.KEY_CODE LOCATION_NO "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (DI.LOCATION_ID = LC.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.UNIT_CHANGE_ID = MR5.KEY_ID) "
      
      TempStr = ""
   ElseIf Ind = 94 Then 'สำหับ report billing019_1
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC "
      SelectStr1 = SelectStr1 & ",MR4.KEY_CODE STOCK_TYPE_CODE ,MR4.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR5.KEY_CODE STOCK_TYPE_SUB_CODE ,MR5.KEY_NAME STOCK_TYPE_SUB_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.STOCK_TYPE = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE_SUB = MR5.KEY_ID) "
      
      TempStr = ""
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO IS NOT NULL ) "
   ElseIf Ind = 95 Then 'สำหับ GetSaleAmountDocumentNo
      SelectStr1 = "SELECT  NM.DOCUMENT_NO,NM.VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.REFER_DESC,NM.CUS_PO "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      TempStr = "WHERE "
   ElseIf Ind = 96 Then       'GenerateComExAmount
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT,SUM(DI.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.GROUP_COM_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = " WHERE NM.APAR_MAS_ID IN  " & APAR_CODE & " AND DI.PART_ITEM_ID IN " & STOCK_NO
      
      APAR_CODE = ""
      STOCK_NO = ""
   ElseIf Ind = 97 Then 'GetAmountForTranSport
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,NM.CUSTOMER_BRANCH,DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",NM.DRIVER_ID,NM.CAR_LICENSE_ID,NM.TRANSPORTOR_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      WhereStr = "WHERE DI.DOC_ITEM_TYPE = 1 "
   ElseIf Ind = 98 Then 'CReportBilling037
      SelectStr1 = "SELECT  DISTINCT NM.APAR_MAS_ID "
      SelectStr1 = SelectStr1 & ",NM.DRIVER_ID,NM.CAR_LICENSE_ID,NM.TRANSPORTOR_ID,NM.CUSTOMER_BRANCH "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE, NM2.SHORT_NAME APAR_NAME "
      
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME DRIVER_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CAR_LICENSE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME TRANSPORTOR_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME CUSTOMER_BRANCH_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.DRIVER_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.CAR_LICENSE_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.TRANSPORTOR_ID = MR3.KEY_ID) "
      
      
      
   ElseIf Ind = 100 Then      'LoadDisTinctBillingDocID
      SelectStr1 = "SELECT  NM.BILLING_DOC_ID,NM.DOCUMENT_SUB_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
   ElseIf Ind = 101 Then      'LoadDisTinctPOID
      SelectStr1 = "SELECT DISTINCT DI.PO_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
   ElseIf Ind = 102 Then      'GetSaleAmountEmpDocType
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_BRANCH IS NOT NULL) " & vbCrLf
   ElseIf Ind = 103 Then 'GetSaleAmountEmpDocTypeBranch
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_BRANCH IS NULL) " & vbCrLf
   
   ElseIf Ind = 104 Then      'GetSaleAmountEmpAparBranchCode
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE APAR_CODE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.CUSTOMER_BRANCH = MR2.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_BRANCH IS NOT NULL) " & vbCrLf
   ElseIf Ind = 105 Then 'GetSaleAmountEmpAparBranchCodeBranch
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE APAR_CODE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE CUSTOMER_BRANCH_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_BRANCH IS NULL) " & vbCrLf
   ElseIf Ind = 106 Then   'GetCountAllCycleDriver
      SelectStr1 = "SELECT DISTINCT NM.TRANSPORT_CYCLE,NM.DRIVER_ID,NM.DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.DRIVER_ID IS NOT NULL) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (NM.TRANSPORT_CYCLE IS NOT NULL) " & vbCrLf
   ElseIf Ind = 107 Then   'GetCountAllCycleTranSportor
      SelectStr1 = "SELECT DISTINCT NM.TRANSPORT_CYCLE,NM.TRANSPORTOR_ID,NM.DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.TRANSPORTOR_ID IS NOT NULL) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (NM.TRANSPORT_CYCLE IS NOT NULL) " & vbCrLf
   ElseIf Ind = 108 Then   'GetSumAmountByDriver
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT,NM.DRIVER_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.STOCK_TYPE = MR4.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.DRIVER_ID IS NOT NULL) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (NM.TRANSPORT_CYCLE IS NOT NULL) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (MR4.PARENT_ID = 74 ) " & vbCrLf 'ID ไข่
   
   ElseIf Ind = 109 Then   'GetSumAmountByTranSportor
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT,NM.TRANSPORTOR_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.STOCK_TYPE = MR4.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.TRANSPORTOR_ID IS NOT NULL) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (NM.TRANSPORT_CYCLE IS NOT NULL) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (MR4.PARENT_ID = 74 ) " & vbCrLf 'ID ไข่
   ElseIf Ind = 110 Then 'GetSaleAmountBranchAparStockCodeDocTypeFreeyyyymm
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_CODE CUSTOMER_BRANCH_CODE,NM.YYYYMM "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 111 Then 'GetSaleAmountBranchAparStockCodeDocTypeFreeExyyyymm
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE CUSTOMER_BRANCH_CODE,NM.YYYYMM "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N' AND PL.BRANCH_ID IS NOT NULL) " & vbCrLf
   ElseIf Ind = 112 Then 'CReportBilling037_1
      SelectStr1 = "SELECT  DISTINCT "
      SelectStr1 = SelectStr1 & " NM.DRIVER_ID,NM.CAR_LICENSE_ID,NM.TRANSPORTOR_ID "
      
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME DRIVER_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CAR_LICENSE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME TRANSPORTOR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.DRIVER_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.CAR_LICENSE_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.TRANSPORTOR_ID = MR3.KEY_ID) "
   ElseIf Ind = 113 Then 'GetAmountForTranSport2
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",NM.DRIVER_ID,NM.CAR_LICENSE_ID,NM.TRANSPORTOR_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      WhereStr = "WHERE DI.DOC_ITEM_TYPE = 1 "
   ElseIf Ind = 114 Then 'GetDistinctForTranSport2
      SelectStr1 = "SELECT  DISTINCT DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SC.PART_DESC "
      SelectStr1 = SelectStr1 & ",SC.UNIT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   ElseIf Ind = 115 Then 'GetDistinctForTranSport
      SelectStr1 = "SELECT  DISTINCT DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SC.PART_DESC "
      SelectStr1 = SelectStr1 & ",SC.UNIT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      '------------------------------------------------------------------------------------------------------------------------------WHERE เฉพาะ ของ ind = 115
   ElseIf Ind = 116 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeDateFree
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_DATE,NM.CUSTOMER_BRANCH "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
      'WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 117 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFree
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_DATE,PL.BRANCH_ID CUSTOMER_BRANCH "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 118 Then 'GetDistinctEmpAparStockCodeDocTypeGroupName
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR6.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",MR7.KEY_NAME CUSTOMER_BRANCH_NAME,NM.CUSTOMER_BRANCH,MR7.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR7 ON (MR7.KEY_ID = NM.CUSTOMER_BRANCH) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (APM.APAR_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 119 Then 'GetDistinctEmpAparStockCodeDocTypeGroupName
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME ,NM2.SHORT_NAME APAR_SHORT_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CUSTOMER_BRANCH_NAME,PL.BRANCH_ID CUSTOMER_BRANCH,MR2.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR3.PARENT_ID = MR4.KEY_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
ElseIf Ind = 120 Then 'GetDistinctEmpAparStockCodeDocTypeGroupName2
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
'      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR6.KEY_NAME APAR_GROUP_NAME , MR8.KEY_NAME STOCK_GROUP_NAME  "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (APM.APAR_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR7 ON (SC.STOCK_TYPE = MR7.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR8 ON (MR7.PARENT_ID = MR8.KEY_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 121 Then 'GetDistinctEmpAparStockCodeDocTypeGroupName2
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
'      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME , MR6.KEY_NAME STOCK_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR3.PARENT_ID = MR4.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 122 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeDateFree2
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      '  ใช้ทำอะไร
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR3.PARENT_ID = MR4.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
      'WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 123 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFree2
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR3.PARENT_ID = MR4.KEY_ID) "
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
ElseIf Ind = 124 Then 'GetDistinctEmpAparDocTypeFree
      SelectStr1 = "SELECT DISTINCT APM.APAR_CODE,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME,MR6.KEY_NAME APAR_GROUP_NAME "
'      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
'      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (APM.APAR_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      WhereStr = "WHERE (APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 125 Then 'GetDistinctStockGroupDocTypeFree
      SelectStr1 = "SELECT DISTINCT MR6.KEY_CODE STOCK_TYPE_CODE ,MR6.KEY_NAME STOCK_GROUP_NAME "
'      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
'      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      WhereStr = "WHERE (APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 126 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM2
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
'      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
'      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
'      SelectStr1 = SelectStr1 & ",NM.YYYYMM "
      SelectStr1 = SelectStr1 & ",MR6.KEY_CODE STOCK_TYPE_CODE  "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) "
   ElseIf Ind = 127 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM2
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
'      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
'      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
'      SelectStr1 = SelectStr1 & ",NM.YYYYMM  "
      SelectStr1 = SelectStr1 & ",MR6.KEY_CODE STOCK_TYPE_CODE  "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 128 Then 'GetDistinctEmpAparDocTypeFree
      SelectStr1 = "SELECT DISTINCT APM.APAR_CODE,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME,MR4.KEY_NAME APAR_GROUP_NAME "
'      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_NAME "
'      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
'      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME APAR_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR3.PARENT_ID = MR4.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "

      WhereStr = "WHERE (APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 130 Then 'LoadNote
      SelectStr1 = "SELECT DISTINCT NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.NOTE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      WhereStr = "WHERE NM.NOTE <> '' "
   ElseIf Ind = 131 Then 'CReportBilling036_10
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_NAME,MR3.KEY_NAME UNIT_CHANGE_NAME,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME,APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 132 Then 'CReportBilling036_10
      SelectStr1 = "SELECT SUM(PL.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME UNIT_NAME,MR3.KEY_NAME UNIT_CHANGE_NAME,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME,APM.APAR_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "

      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (PL.BRANCH_ID = MR4.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR4.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR4.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "

      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 133 Then 'CReportBilling036_11
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME,APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME CUSTOMER_BRANCH_NAME,NM.CUSTOMER_BRANCH,MR5.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = NM.CUSTOMER_BRANCH) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 134 Then 'CReportBilling036_11
      SelectStr1 = "SELECT SUM(PL.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME,APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_NAME CUSTOMER_BRANCH_NAME,PL.BRANCH_ID CUSTOMER_BRANCH,MR4.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "

      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (PL.BRANCH_ID = MR4.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR4.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR4.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "

      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 135 Then 'GetDistinctEmpAparStockCodeDocTypeFreeNotBranch
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_MAS_GROUP_CODE,APM.APAR_MAS_GROUP_NAME,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
     
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 136 Then 'GetDistinctEmpAparStockCodeDocTypeFreeNotBranch
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,APM.APAR_MAS_GROUP_CODE,APM.APAR_MAS_GROUP_NAME,APM.APAR_NAME,NM2.SHORT_NAME APAR_SHORT_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL OR APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 137 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeDateFreeNotBranch
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_DATE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "

      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
   ElseIf Ind = 138 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFreeBranch
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_DATE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 139 Then 'GetDistinctStockGroupDocTypeFree
      SelectStr1 = "SELECT DISTINCT MR6.KEY_CODE STOCK_TYPE_CODE ,MR6.KEY_NAME STOCK_GROUP_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
      WhereStr = "WHERE (APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 140 Then 'GetDistinctStockCodeDocTypeFree2
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.PART_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      WhereStr = "WHERE (APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 141 Then 'GetDistinctStockCodeDocTypeFree2
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.PART_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      WhereStr = "WHERE (APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 142 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM3
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "
'      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
'      SelectStr1 = SelectStr1 & ",NM.YYYYMM "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) "
   ElseIf Ind = 143 Then 'GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM3
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE,APM.APAR_CODE "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
    ElseIf Ind = 144 Then 'LoadNote
      SelectStr1 = "SELECT DISTINCT NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.NOTE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "

      WhereStr = "WHERE NM.NOTE <> '' "
   ElseIf Ind = 145 Then 'reportbillingdoc012_1
      SelectStr1 = "SELECT SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT  "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.PAY_AMOUNT)PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
  
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE,NM.BILLING_DOC_ID "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE SALE_CODE,NM4.LONG_NAME SALE_NAME,NM4.LAST_NAME SALE_LAST_NAME "
'      SelectStr1 = SelectStr1 & ",NM.DUE_DATE,NM3.DOCUMENT_NO DOC_ID_BILLS_NO,NM3.DUE_DATE DOC_ID_BILLS_DUE_DATE,APM.APAR_TYPE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE,APM.APAR_TYPE,MR2.SHORT_CODE LOCATION_CODE "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME CUSTOMER_BRANCH_NAME,PL.BRANCH_ID CUSTOMER_BRANCH,MR3.KEY_CODE CUSTOMER_BRANCH_CODE "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
       FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (PL.BRANCH_ID = MR3.KEY_ID) "
       FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR3.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
       
       
       'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR3.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN1.NAME_ID = NM4.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP1 ON (NM.BILLING_DOC_ID = RCP1.DOC_ID_BILLS   ) "
'      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM3 ON (RCP1.BILLING_DOC_ID = NM3.BILLING_DOC_ID ) "
   ElseIf Ind = 146 Then 'reportbillingdoc012_1
      SelectStr1 = "SELECT SUM(NM.TOTAL_PRICE) TOTAL_PRICE,SUM(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT  "
      SelectStr1 = SelectStr1 & ",SUM(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.PAY_AMOUNT) PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE,NM.BILLING_DOC_ID "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE SALE_CODE,NM4.LONG_NAME SALE_NAME,NM4.LAST_NAME SALE_LAST_NAME "
'      SelectStr1 = SelectStr1 & ",NM.DUE_DATE,NM3.DOCUMENT_NO DOC_ID_BILLS_NO,NM3.DUE_DATE DOC_ID_BILLS_DUE_DATE,APM.APAR_TYPE "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE,APM.APAR_TYPE,MR2.SHORT_CODE LOCATION_CODE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME CUSTOMER_BRANCH_NAME,NM.CUSTOMER_BRANCH,MR5.KEY_CODE CUSTOMER_BRANCH_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN1.NAME_ID = NM4.NAME_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP1 ON (NM.BILLING_DOC_ID = RCP1.DOC_ID_BILLS   ) "
'      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM3 ON (RCP1.BILLING_DOC_ID = NM3.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = NM.CUSTOMER_BRANCH) "

'      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL) "
      Dim TempEMPCODE As String
      TempEMPCODE = "1-ซ027B"
      WhereStr = "WHERE (APM.APAR_CODE <> " & "'" & TempEMPCODE & "'" & " ) "
   ElseIf Ind = 147 Then 'GetDistinctEmpDocTypeFree
      SelectStr1 = "SELECT DISTINCT  EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,EM.JOINT_CODE   SALE_JOINT_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = NM.CUSTOMER_BRANCH) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL AND APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 148 Then 'GetDistinctEmpDocTypeFree
      SelectStr1 = "SELECT DISTINCT EM.EMP_CODE SALE_CODE,EM.NOT_SHOW_RETURN,EM.JOINT_CODE   SALE_JOINT_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      
      
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL AND APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 149 Then 'GetSaleAmountEmp
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE   SALE_JOINT_CODE,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
   ElseIf Ind = 150 Then 'GetSaleAmountEmp
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE   SALE_JOINT_CODE,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID ) "
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
   ElseIf Ind = 151 Then 'GetDistinctEmpGroupStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR8.KEY_CODE STOCK_TYPE_CODE,MR8.KEY_NAME STOCK_GROUP_NAME  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR7 ON (SC.STOCK_TYPE = MR7.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR8 ON (MR7.PARENT_ID = MR8.KEY_ID) "

      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL AND APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 152 Then 'GetDistinctEmpGroupStockCodeDocTypeFree
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",MR8.KEY_CODE STOCK_TYPE_CODE,MR8.KEY_NAME STOCK_GROUP_NAME  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "

      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR7 ON (SC.STOCK_TYPE = MR7.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR8 ON (MR7.PARENT_ID = MR8.KEY_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL AND APM.APAR_CODE IS NOT NULL) "
   ElseIf Ind = 153 Then 'GetSaleAmountEmpStockCodeDocTypeFreeDate
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_DATE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "

      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
      WhereStr = WhereStr & " AND (APM.APAR_CODE != '" & "ค-001" & "')"
   ElseIf Ind = 154 Then 'GetSaleAmountEmpStockCodeDocTypeFreeExDate
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE "
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_DATE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
      WhereStr = WhereStr & " AND (APM.APAR_CODE != '" & "ค-001" & "')"
   ElseIf Ind = 155 Then 'GetSaleAmountStockCodeDocTypeForFAndB
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      WhereStr = " WHERE (APM.APAR_CODE != '" & "ค-001" & "')"
   ElseIf Ind = 156 Then 'สำหรับ GetDistinctStockCodeForFAndB
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT,SC.JOINT_CODE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME STOCK_GROUP_NAME,MR5.KEY_CODE STOCK_GROUP_CODE "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = MR2.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
      
   ElseIf Ind = 157 Then
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_DATE,MR5.KEY_CODE STOCK_GROUP_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = MR2.PARENT_ID) "
   ElseIf Ind = 158 Then 'GetDistinctEmpStockCodeDocTypeFreeForFAndB
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.JOINT_CODE,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE SALE_JOINT_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
      
   ElseIf Ind = 159 Then 'GetDistinctEmpDocTypeFreeForFAndB
      SelectStr1 = "SELECT DISTINCT  EM.EMP_CODE SALE_CODE,EM.JOINT_CODE   SALE_JOINT_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = NM.CUSTOMER_BRANCH) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = "WHERE (EM.EMP_CODE IS NOT NULL AND APM.APAR_CODE IS NOT NULL) "
      
   ElseIf Ind = 160 Then 'GetSaleAmountEmpForFAndB
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE   SALE_JOINT_CODE,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID IS NULL) " & vbCrLf
      
   ElseIf Ind = 161 Then 'GetSaleAmountEmpForFAndB
      SelectStr1 = "SELECT SUM(PL.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.TOTAL_PRICE) TOTAL_PRICE,SUM(PL.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.JOINT_CODE   SALE_JOINT_CODE,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID ) "
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = 'N') " & vbCrLf
      
   ElseIf Ind = 162 Then 'reportbillingdoc012_2
      SelectStr1 = "SELECT MAX(NM.TOTAL_PRICE) TOTAL_PRICE,MAX(NM.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.EXT_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.VAT_AMOUNT) VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MAX(NM.PAY_AMOUNT) PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ",RCP.PAID_AMOUNT PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",RCP.PAID_DISCOUNT PAID_DISCOUNT "                     'ส่วนลดรับชำระ
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE SALE_CODE,NM4.LONG_NAME SALE_NAME "
      SelectStr1 = SelectStr1 & ",NM.DUE_DATE,NM3.DOCUMENT_NO DOC_ID_BILLS_NO,NM3.DUE_DATE DOC_ID_BILLS_DUE_DATE,APM.APAR_TYPE "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID ) "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (NM.BILLING_DOC_ID = RCP.DOC_ID_RCP   ) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM1 ON (RCP.BILLING_DOC_ID = NM1.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN1.NAME_ID = NM4.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP1 ON (NM.BILLING_DOC_ID = RCP1.DOC_ID_BILLS   ) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC NM3 ON (RCP1.BILLING_DOC_ID = NM3.BILLING_DOC_ID ) "
      TempStr = ""
   ElseIf Ind = 163 Then 'reportbillingdocD002_1
      SelectStr1 = "SELECT NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE,NM.BILLING_DOC_ID "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.DUE_DATE,EM.EMP_CODE SALE_CODE "
      'SelectStr1 = SelecStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      TempStr = ""
   ElseIf Ind = 164 Then 'reportbillingdocD002_2
      SelectStr1 = "SELECT NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.VAT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.PAY_AMOUNT "            ' เพิ่มหนี้ ลดหนี้
      
      SelectStr1 = SelectStr1 & ", NM.DOCUMENT_TYPE,NM.BILLING_DOC_ID "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.DUE_DATE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.SALE_BY = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      TempStr = ""
   ElseIf Ind = 165 Then 'สำหรับ GetDistinctStockCode_Y
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT,SC.JOINT_CODE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME STOCK_GROUP_NAME,MR5.KEY_CODE STOCK_GROUP_CODE "
      SelectStr1 = SelectStr1 & ",APM.CONSIGNMENT_FLAG APM_CONSIGNMENT_FLAG  "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = MR2.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
 ElseIf Ind = 166 Then 'สำหรับ GetDistinctStockCodeForFAndB
      SelectStr1 = "SELECT DISTINCT SC.STOCK_NO,SC.STOCK_DESC,SC.STOCK_TYPE,MR2.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",SC.CHK_STD_COST,SC.COST_PER_AMOUNT,SC.JOINT_CODE "
      SelectStr1 = SelectStr1 & ",MR5.KEY_NAME STOCK_GROUP_NAME,MR5.KEY_CODE STOCK_GROUP_CODE "
     SelectStr1 = SelectStr1 & ",APM.CONSIGNMENT_FLAG APM_CONSIGNMENT_FLAG  "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.STOCK_TYPE = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (APM.APAR_TYPE = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR4.KEY_ID = MR3.PARENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (MR5.KEY_ID = MR2.PARENT_ID) "
      TempStr = ""
      
      WhereStr = " WHERE DI.PART_ITEM_ID IS NOT NULL "
      
   ElseIf Ind = 167 Then      'GetSaleAmountCustomerStockGroup
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID,APM.APAR_TYPE,MR4.PARENT_ID STOCK_GROUP "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (SC.STOCK_TYPE = MR4.KEY_ID) "
   ElseIf Ind = 168 Then      'GetSaleAmountByCustomer
      SelectStr1 = "SELECT SUM(DI.ITEM_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(DI.EXTRA_DISCOUNT_AMOUNT) EXT_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.APAR_MAS_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILLING_DOC NM "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (NM.BILLING_DOC_ID = DI.BILLING_DOC_ID ) "
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
      
   If APAR_MAS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.APAR_MAS_ID = " & APAR_MAS_ID & ") " & vbCrLf
   End If
   
   'WhereStr = " WHERE NM.TRANSPORTOR_ID IS NOT NULL "
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (DI.LOCATION_ID = " & LOCATION_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_PACK > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BILLING_DOC_PACK = " & BILLING_DOC_PACK & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If APAR_IND > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_IND = " & APAR_IND & ") " & vbCrLf
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
      End If
   End If
   
   If Len(Trim(DOCUMENT_TYPE_SET)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_TYPE IN " & ChangeQuote(Trim(DOCUMENT_TYPE_SET)) & ")"
   End If
      
   If Len(Trim(APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
      End If
   End If
   
   If Len(Trim(SHORT_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR2.SHORT_CODE = '" & ChangeQuote(Trim(SHORT_CODE)) & "'"
         If SHORT_CODE = 2 Then
             WhereStr = WhereStr & " OR MR2.SHORT_CODE is null )"          ' ไม่ใช่ห้าง หรือ ที่ยังไม่ระบุ
         Else
              WhereStr = WhereStr & ")"
         End If
   End If
   
   If Len(Trim(FROM_APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_CODE >= '" & ChangeQuote(Trim(FROM_APAR_CODE)) & "')"
   End If
   
   If Len(Trim(TO_APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_CODE <= '" & ChangeQuote(Trim(TO_APAR_CODE)) & "')"
   End If
   
   If Len(Trim(SALE_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(SALE_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(EM.EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (EM.EMP_CODE = '" & ChangeQuote(Trim(SALE_CODE)) & "')"
      End If
   End If
   
   If Len(Trim(FROM_SALE_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.EMP_CODE >= '" & ChangeQuote(Trim(FROM_SALE_CODE)) & "')"
   End If
   
   If Len(Trim(TO_SALE_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.EMP_CODE <= '" & ChangeQuote(Trim(TO_SALE_CODE)) & "')"
   End If
   
   If Len(Trim(FROM_LOCATION_NO)) > 0 Then
      Call AddSearchLocationNo(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (LC.KEY_CODE >= '" & ChangeQuote(Trim(FROM_LOCATION_NO)) & "')"
   End If
   
   If Len(Trim(TO_LOCATION_NO)) > 0 Then
      Call AddSearchLocationNo(FromStr)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (LC.KEY_CODE <= '" & ChangeQuote(Trim(TO_LOCATION_NO)) & "')"
   End If
   
   If Len(Trim(APAR_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APM.APAR_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APM.APAR_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
      End If
   End If
   
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If FROM_DUE_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DUE_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DUE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DUE_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DUE_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DUE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 1 Then
      If Len(Trim(STOCK_NO)) > 0 Then
         If Len(WhereStr) > 0 Then
            TempStr = "AND "
         Else
            TempStr = "WHERE "
         End If
         If WildCard(STOCK_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM DOC_ITEM DI LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) AND (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)"
         Else
            WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM DOC_ITEM DI LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (DI.BILLING_DOC_ID = NM.BILLING_DOC_ID) AND SC.STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "') > 0)"
         End If
      End If
   Else
      If Len(Trim(STOCK_NO)) > 0 Then
         If Len(WhereStr) > 0 Then
            TempStr = "AND "
         Else
            TempStr = "WHERE "
         End If
         If WildCard(STOCK_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & TempStr & " (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & TempStr & " (SC.STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "')"
         End If
      End If
   End If
   
   If Trim(FROM_STOCK_NO) <> "" Then
      Call AddSearchStockCode(FromStr)
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(TO_STOCK_NO) <> "" Then
      Call AddSearchStockCode(FromStr)
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      End If
   End If
      
   If Len(Trim(DOC_ID_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(DOC_ID_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM RCPCNDN_ITEM RCP LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = RCP.DOC_ID) WHERE (RCP.BILLING_DOC_ID = NM.BILLING_DOC_ID) AND (SUBSTR(BD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "') > 0)"
      Else
         WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM RCPCNDN_ITEM RCP LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = RCP.DOC_ID) WHERE (RCP.BILLING_DOC_ID = NM.BILLING_DOC_ID) AND (BD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOC_ID_NO)) & "')) > 0)"
      End If
   End If
      
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CANCEL_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.CANCEL_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
      
   
   If Ind = 3 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME,NM.DOCUMENT_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = " GROUP BY NM.SALE_BY,SC.GROUP_COM_ID "
   ElseIf Ind = 9 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME,DI.UNIT_MULTIPLE "
   ElseIf Ind = 10 Then
      GroupByStr = " GROUP BY NM.DUE_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME,DI.UNIT_MULTIPLE "
   ElseIf Ind = 12 Then
      GroupByStr = " GROUP BY NM.APAR_MAS_ID,NM.DOCUMENT_TYPE "
   ElseIf Ind = 13 Then
      'GroupByStr = " GROUP BY NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM.BILLING_DOC_ID,MR2.KEY_CODE,MR2.KEY_NAME,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE "
   ElseIf Ind = 14 Then
      GroupByStr = " GROUP BY NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM.DUE_DATE,APM.APAR_TYPE "    ',NM3.DOCUMENT_NO ,NM3.DUE_DATE
      GroupByStr = GroupByStr & ",NM.TOTAL_PRICE,NM.DISCOUNT_AMOUNT,NM.EXT_DISCOUNT_AMOUNT,NM.VAT_AMOUNT,NM.PAY_AMOUNT "
   ElseIf Ind = 20 Then
      GroupByStr = " GROUP BY APM.APAR_TYPE,NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM.DUE_DATE,NM3.DOCUMENT_NO ,NM3.DUE_DATE, RCP.PAID_AMOUNT, RCP.PAID_DISCOUNT "
   ElseIf Ind = 21 Then
      GroupByStr = " GROUP BY NM.APAR_MAS_ID,NM.YYYYMM,NM.DOCUMENT_TYPE "
   ElseIf Ind = 22 Then
      GroupByStr = " GROUP BY MR2.KEY_NAME,SC.STOCK_NO,DI.PART_ITEM_ID,SC.STOCK_DESC,NM.DOCUMENT_TYPE,DI.YYYYMM,MR6.KEY_NAME "
   ElseIf Ind = 23 Then
      GroupByStr = " GROUP BY NM.BILLING_DOC_ID,NM.DOCUMENT_TYPE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,NM.REFER_TEXT,NM.CUS_PO "
   ElseIf Ind = 24 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.YYYYMM,NM.DOCUMENT_TYPE  "
   ElseIf Ind = 25 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME ,MR3.KEY_NAME,NM.DOCUMENT_TYPE,SC.UNIT_AMOUNT  "
   ElseIf Ind = 27 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.YYYYMM,NM.DOCUMENT_TYPE,APM.APAR_CODE  "
   ElseIf Ind = 29 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME ,MR3.KEY_NAME,NM.DOCUMENT_TYPE,DI.FREE_FLAG  "
   ElseIf Ind = 36 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME ,MR3.KEY_NAME,NM.DOCUMENT_TYPE,DI.FREE_FLAG,APM.APAR_CODE  "
   ElseIf Ind = 37 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,EM.EMP_CODE,MR4.KEY_CODE "
   ElseIf Ind = 38 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE,MR4.KEY_CODE  "
   ElseIf Ind = 39 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE "
   ElseIf Ind = 42 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,EM.EMP_CODE,MR2.KEY_CODE "
   ElseIf Ind = 45 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE,MR4.KEY_CODE "
   ElseIf Ind = 46 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE,MR2.KEY_CODE "
   ElseIf Ind = 50 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,MR4.KEY_CODE "
   ElseIf Ind = 51 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,MR2.KEY_CODE "
   ElseIf Ind = 54 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,MR4.KEY_CODE "
   ElseIf Ind = 55 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,MR2.KEY_CODE "
   ElseIf Ind = 58 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,EM.EMP_CODE "
   ElseIf Ind = 59 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,EM.EMP_CODE "
   ElseIf Ind = 62 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE "
   ElseIf Ind = 63 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE "
   ElseIf Ind = 66 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE "
   ElseIf Ind = 67 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE "
   ElseIf Ind = 75 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,MR5.KEY_NAME,NM.DOCUMENT_TYPE  "
   ElseIf Ind = 77 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.JOINT_CODE,NM.DOCUMENT_TYPE "
   ElseIf Ind = 78 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,MR4.KEY_NAME,SC.CHK_STD_COST,SC.COST_PER_AMOUNT,NM.DOCUMENT_TYPE   "
   ElseIf Ind = 84 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,MR4.KEY_NAME,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE   "
   ElseIf Ind = 85 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE "
   ElseIf Ind = 86 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,EM.EMP_CODE,NM.YYYYMM,NM.CUSTOMER_BRANCH  "
   ElseIf Ind = 87 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,EM.EMP_CODE,NM.YYYYMM,PL.BRANCH_ID  "
   ElseIf Ind = 90 Then
      GroupByStr = " GROUP BY NM.APAR_MAS_ID,NM.DOCUMENT_TYPE "
   ElseIf Ind = 91 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,MR5.KEY_NAME,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE   "
   ElseIf Ind = 95 Then
      GroupByStr = " GROUP BY NM.VAT_AMOUNT,NM.DOCUMENT_NO,NM.REFER_DESC,NM.CUS_PO   "
   ElseIf Ind = 96 Then
      GroupByStr = " GROUP BY SC.GROUP_COM_ID   "
   ElseIf Ind = 97 Then
      GroupByStr = " GROUP BY NM.APAR_MAS_ID,NM.CUSTOMER_BRANCH,DI.PART_ITEM_ID,NM.DRIVER_ID,NM.CAR_LICENSE_ID,NM.TRANSPORTOR_ID  "
   ElseIf Ind = 102 Or Ind = 103 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_TYPE,EM.EMP_CODE "
   ElseIf Ind = 104 Or Ind = 105 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_TYPE,EM.EMP_CODE,APM.APAR_CODE,MR2.KEY_CODE "
   ElseIf Ind = 108 Then
      GroupByStr = " GROUP BY NM.DRIVER_ID "
   ElseIf Ind = 109 Then
      GroupByStr = " GROUP BY NM.TRANSPORTOR_ID "
   ElseIf Ind = 110 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,MR4.KEY_CODE,NM.YYYYMM "
   ElseIf Ind = 111 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE,APM.APAR_CODE,MR2.KEY_CODE,NM.YYYYMM "
   ElseIf Ind = 113 Then
      GroupByStr = " GROUP BY DI.PART_ITEM_ID,NM.DRIVER_ID,NM.CAR_LICENSE_ID,NM.TRANSPORTOR_ID  "
   ElseIf Ind = 116 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,APM.APAR_CODE,NM.CUSTOMER_BRANCH,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 117 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,APM.APAR_CODE,PL.BRANCH_ID,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 122 Then
      GroupByStr = " GROUP BY MR4.KEY_NAME,SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 123 Then
      GroupByStr = " GROUP BY MR4.KEY_NAME,SC.STOCK_NO,NM.DOCUMENT_TYPE,EM.EMP_CODE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 126 Then
      GroupByStr = " GROUP BY APM.APAR_CODE,MR6.KEY_CODE,NM.DOCUMENT_TYPE  "
   ElseIf Ind = 127 Then
      GroupByStr = " GROUP BY APM.APAR_CODE,MR6.KEY_CODE,NM.DOCUMENT_TYPE  "
   ElseIf Ind = 131 Or Ind = 132 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME ,MR3.KEY_NAME,NM.DOCUMENT_TYPE,SC.UNIT_AMOUNT,EM.EMP_CODE ,NM3.LONG_NAME ,NM3.LAST_NAME,APM.APAR_CODE   "
   ElseIf Ind = 133 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_TYPE,EM.EMP_CODE ,NM3.LONG_NAME ,NM3.LAST_NAME,APM.APAR_CODE,NM2.LONG_NAME,MR5.KEY_NAME,NM.CUSTOMER_BRANCH,MR5.KEY_CODE   "
   ElseIf Ind = 134 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_TYPE,EM.EMP_CODE ,NM3.LONG_NAME ,NM3.LAST_NAME,APM.APAR_CODE,NM2.LONG_NAME,MR4.KEY_NAME,PL.BRANCH_ID,MR4.KEY_CODE   "
   ElseIf Ind = 137 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,APM.APAR_CODE,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 138 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,APM.APAR_CODE,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 142 Or Ind = 143 Then
      GroupByStr = " GROUP BY APM.APAR_CODE,SC.STOCK_NO,NM.DOCUMENT_TYPE  "
   ElseIf Ind = 145 Then
'      GroupByStr = " GROUP BY APM.APAR_TYPE,NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM.DUE_DATE,NM3.DOCUMENT_NO ,NM3.DUE_DATE,MR3.KEY_NAME,PL.BRANCH_ID,MR3.KEY_CODE,NM4.LONG_NAME,NM4.LAST_NAME,NM.BILLING_DOC_ID "
      GroupByStr = " GROUP BY APM.APAR_TYPE,NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM.DUE_DATE,MR3.KEY_NAME,PL.BRANCH_ID,MR3.KEY_CODE,NM4.LONG_NAME,NM4.LAST_NAME,NM.BILLING_DOC_ID,MR2.SHORT_CODE "
   ElseIf Ind = 146 Then
'      GroupByStr = " GROUP BY APM.APAR_TYPE,NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM.DUE_DATE,NM3.DOCUMENT_NO ,NM3.DUE_DATE,MR5.KEY_NAME,NM.CUSTOMER_BRANCH,MR5.KEY_CODE,NM4.LONG_NAME,NM4.LAST_NAME,NM.BILLING_DOC_ID "
      GroupByStr = " GROUP BY APM.APAR_TYPE,NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM.DUE_DATE,MR5.KEY_NAME,NM.CUSTOMER_BRANCH,MR5.KEY_CODE,NM4.LONG_NAME,NM4.LAST_NAME,NM.BILLING_DOC_ID,MR2.SHORT_CODE "
   ElseIf Ind = 149 Or Ind = 150 Or Ind = 160 Or Ind = 161 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,EM.JOINT_CODE,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE  "
   ElseIf Ind = 153 Or Ind = 154 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 155 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,NM.DOCUMENT_TYPE "
   ElseIf Ind = 157 Then
      GroupByStr = " GROUP BY MR5.KEY_CODE,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE  "
   ElseIf Ind = 162 Then
      GroupByStr = " GROUP BY APM.APAR_TYPE,NM.APAR_MAS_ID,NM.DOCUMENT_TYPE,NM2.LONG_NAME,NM2.LAST_NAME,APM.APAR_CODE,NM.DOCUMENT_NO,NM.DOCUMENT_DATE,EM.EMP_CODE,NM4.LONG_NAME,NM.DUE_DATE,NM3.DOCUMENT_NO ,NM3.DUE_DATE , RCP.PAID_AMOUNT , RCP.PAID_DISCOUNT "
   ElseIf Ind = 167 Then
      GroupByStr = " GROUP BY NM.DOCUMENT_TYPE,NM.APAR_MAS_ID,APM.APAR_TYPE,MR4.PARENT_ID "
   ElseIf Ind = 168 Then
      GroupByStr = " GROUP BY NM.APAR_MAS_ID "
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 9999 Then
         'ไม่ต้องการให้ order เพื่อความเร็ว
      ElseIf ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 3 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY SC.STOCK_NO,NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY SC.STOCK_NO,NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY SC.STOCK_NO,NM.DOCUMENT_DATE,NM.TOTAL_AMOUNT "
      Else
         OrderByStr = "ORDER BY SC.STOCK_NO,NM.DOCUMENT_DATE,NM.DOCUMENT_TYPE "
      End If
   ElseIf Ind = 4 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE,NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APM.APAR_CODE,NM.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE,NM.DOCUMENT_TYPE ASC,NM.DOCUMENT_DATE,NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY NM.SALE_BY "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY NM.SALE_BY "
   ElseIf Ind = 7 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE,NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 9 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      End If
   ElseIf Ind = 10 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DUE_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DUE_DATE "
      Else
         OrderByStr = "ORDER BY NM.DUE_DATE "
      End If
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE ASC,APM.APAR_CODE "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE ASC,APM.APAR_CODE ,NM.DOCUMENT_NO "
   ElseIf Ind = 14 Then
      OrderByStr = "ORDER BY APM.APAR_TYPE ASC,NM.APAR_MAS_ID,NM.DUE_DATE "
   ElseIf Ind = 15 Then
      OrderByStr = "ORDER BY NM.APAR_MAS_ID,NM.DOCUMENT_NO,SC.STOCK_NO "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY NM.CUSTOMER_BRANCH,EM.EMP_CODE,NM.DOCUMENT_NO,SC.STOCK_NO "
   ElseIf Ind = 18 Then
      OrderByStr = "ORDER BY SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE,NM.DOCUMENT_NO "
   ElseIf Ind = 19 Then
      OrderByStr = "ORDER BY APM.APAR_CODE,SC.STOCK_NO,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE,NM.DOCUMENT_NO "
   ElseIf Ind = 20 Then
      OrderByStr = "ORDER BY APM.APAR_TYPE ASC,NM.APAR_MAS_ID,NM.DUE_DATE "
   ElseIf Ind = 23 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_NO,NM.DOCUMENT_DATE "
   ElseIf Ind = 26 Then
      OrderByStr = "ORDER BY APM.APAR_CODE ASC,SC.STOCK_NO "
   ElseIf Ind = 28 Then
      OrderByStr = "ORDER BY APM.APAR_CODE ASC,NM.DOCUMENT_DATE ASC,NM.DOCUMENT_NO ASC,CQ.CHEQUE_NO,NM1.DOCUMENT_NO "
   ElseIf Ind = 30 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_NO "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY NM.DUE_DATE "
      ElseIf ORDER_BY = 4 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE,NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 31 Or Ind = 32 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_NO "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY NM.DUE_DATE "
      ElseIf ORDER_BY = 4 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE,NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 34 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE ,SC.STOCK_NO "
   ElseIf Ind = 35 Then
      OrderByStr = "ORDER BY APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 40 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE,EM.EMP_CODE,APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 41 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE,EM.EMP_CODE,APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 43 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE,EM.EMP_CODE,SC.STOCK_NO "
   ElseIf Ind = 44 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE,EM.EMP_CODE,SC.STOCK_NO "
   ElseIf Ind = 47 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 48 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE,APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 49 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE,APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 52 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 53 Then
      OrderByStr = "ORDER BY MR2.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 56 Then
      If ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.EMP_CODE,SC.STOCK_NO,APM.APAR_CODE,NM.CUSTOMER_BRANCH "
      Else
         OrderByStr = "ORDER BY EM.EMP_CODE,APM.APAR_CODE,NM.CUSTOMER_BRANCH,SC.STOCK_NO "
      End If
   ElseIf Ind = 57 Then
      If ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.EMP_CODE,SC.STOCK_NO,APM.APAR_CODE,PL.BRANCH_ID "
      Else
         OrderByStr = "ORDER BY EM.EMP_CODE,APM.APAR_CODE,PL.BRANCH_ID,SC.STOCK_NO "
      End If
   ElseIf Ind = 60 Then
      OrderByStr = "ORDER BY EM.EMP_CODE ASC,SC.STOCK_NO "
   ElseIf Ind = 61 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,SC.STOCK_NO "
   ElseIf Ind = 64 Then
      OrderByStr = "ORDER BY APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 65 Then
      OrderByStr = "ORDER BY APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 69 Then
      OrderByStr = "ORDER BY APM.APAR_CODE ASC,NM.DOCUMENT_DATE ASC,NM.DOCUMENT_NO ASC,NM1.DOCUMENT_DATE ASC ,NM1.DOCUMENT_NO "
   ElseIf Ind = 71 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_DATE ASC,NM.DOCUMENT_NO ASC,NM1.DOCUMENT_DATE ASC ,NM1.DOCUMENT_NO "
   ElseIf Ind = 73 Then
      OrderByStr = "ORDER BY EM.EMP_CODE ASC,NM.DOCUMENT_DATE ASC,NM.DOCUMENT_NO ASC,NM1.DOCUMENT_DATE ASC ,NM1.DOCUMENT_NO "
   ElseIf Ind = 74 Then
      If ORDER_TYPE = 2 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE DESC,NM.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_DATE ASC,NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 76 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME,NM.DOCUMENT_TYPE,SC.STOCK_TYPE,SC.STOCK_NO "
   ElseIf Ind = 79 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME ASC,MR3.KEY_NAME ASC,NM.DOCUMENT_TYPE,SC.STOCK_NO "
   ElseIf Ind = 80 Then
      OrderByStr = "ORDER BY MR5.KEY_CODE,MR2.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 83 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME,NM.DOCUMENT_TYPE ,MR3.KEY_NAME,SC.STOCK_TYPE,SC.STOCK_NO "
   ElseIf Ind = 88 Or Ind = 89 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_NO "
   ElseIf Ind = 92 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME,NM.DOCUMENT_TYPE ,SC.STOCK_TYPE,SC.STOCK_NO "
   ElseIf Ind = 93 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_DATE ASC,NM.DOCUMENT_NO "
   ElseIf Ind = 94 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE ,MR5.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 98 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE ,MR2.KEY_CODE,MR3.KEY_CODE,APM.APAR_CODE,MR4.KEY_CODE "
   ElseIf Ind = 99 Or Ind = 114 Or Ind = 115 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 100 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.DOCUMENT_SUB_TYPE,NM.DOCUMENT_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.DRIVER_ID DESC,NM.TRANSPORTOR_ID "
      Else
         OrderByStr = "ORDER BY NM.DOCUMENT_NO "
      End If
   ElseIf Ind = 112 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE ,MR2.KEY_CODE,MR3.KEY_CODE  "
   ElseIf Ind = 118 Then
      OrderByStr = "ORDER BY MR6.KEY_NAME,EM.EMP_CODE,APM.APAR_CODE,NM.CUSTOMER_BRANCH,SC.STOCK_NO "
   ElseIf Ind = 119 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME,EM.EMP_CODE,APM.APAR_CODE,PL.BRANCH_ID,SC.STOCK_NO "
   ElseIf Ind = 120 Then
      OrderByStr = "ORDER BY MR6.KEY_NAME,EM.EMP_CODE,MR8.KEY_NAME,SC.STOCK_NO "
   ElseIf Ind = 121 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME,EM.EMP_CODE,MR6.KEY_NAME,SC.STOCK_NO "
   ElseIf Ind = 124 Then
      OrderByStr = "ORDER BY MR6.KEY_NAME,APM.APAR_CODE "
   ElseIf Ind = 125 Or Ind = 139 Then
      OrderByStr = "ORDER BY MR6.KEY_CODE "
   ElseIf Ind = 128 Then
      OrderByStr = "ORDER BY MR4.KEY_NAME,APM.APAR_CODE "
   ElseIf Ind = 129 Then
      OrderByStr = "ORDER BY APM.APAR_CODE,NM.DOCUMENT_TYPE,NM.DOCUMENT_DATE,NM.DOCUMENT_NO,SC.STOCK_NO "
   ElseIf Ind = 130 Or Ind = 144 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_NO,NM.DOCUMENT_DATE "
   ElseIf Ind = 135 Or Ind = 136 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,APM.APAR_MAS_GROUP_CODE,APM.APAR_CODE,SC.STOCK_NO "
   ElseIf Ind = 140 Or Ind = 141 Then
      OrderByStr = "ORDER BY SC.STOCK_NO,SC.STOCK_DESC,SC.PART_DESC "
   ElseIf Ind = 145 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,APM.APAR_CODE,APM.APAR_TYPE ASC,PL.BRANCH_ID,MR3.KEY_CODE,NM.DUE_DATE "
   ElseIf Ind = 146 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,APM.APAR_CODE,APM.APAR_TYPE ASC,NM.CUSTOMER_BRANCH,MR5.KEY_CODE,NM.DUE_DATE "
   ElseIf Ind = 147 Or Ind = 148 Or Ind = 159 Then
      OrderByStr = "ORDER BY EM.EMP_CODE "
   ElseIf Ind = 151 Or Ind = 152 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,MR8.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 156 Then
      OrderByStr = "ORDER BY MR5.KEY_CODE,MR2.KEY_CODE,SC.STOCK_NO "
   ElseIf Ind = 158 Then
      OrderByStr = "ORDER BY EM.JOINT_CODE ASC,SC.JOINT_CODE "
   ElseIf Ind = 162 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,NM2.LAST_NAME,APM.APAR_TYPE ASC,NM.APAR_MAS_ID,NM.DUE_DATE "
   ElseIf Ind = 163 Then
       OrderByStr = "ORDER BY APM.APAR_CODE ASC,NM.DOCUMENT_NO "
  ElseIf Ind = 164 Then
        OrderByStr = "ORDER BY EM.EMP_CODE ASC, APM.APAR_CODE ASC,NM.DOCUMENT_NO "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
'   Debug.Print SQL1
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      NOTE = NVLS(Rs("NOTE"), "")
      BILLING_ADDRESS_ID = NVLI(Rs("BILLING_ADDRESS_ID"), -1)
      ENTERPRISE_ADDRESS_ID = NVLI(Rs("ENTERPRISE_ADDRESS_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      WH_PERCENT = NVLD(Rs("WH_PERCENT"), 0)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      CREDIT = NVLD(Rs("CREDIT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      SUBTRACT_AMOUNT = NVLD(Rs("SUBTRACT_AMOUNT"), 0)
      ADDITION_AMOUNT = NVLD(Rs("ADDITION_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      BRANCH_ADDRESS = NVLI(Rs("BRANCH_ADDRESS"), -1)
      DOCUMENT_SUB_TYPE = NVLI(Rs("DOCUMENT_SUB_TYPE"), -1)
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      RCP_CASH_TRAN = NVLD(Rs("RCP_CASH_TRAN"), 0)
      RCP_DIF = NVLD(Rs("RCP_DIF"), 0)
      CASH_PMT = NVLD(Rs("CASH_PMT"), 0)
      CHEQUE_PMT = NVLD(Rs("CHEQUE_PMT"), 0)
      BANKTRF_PMT = NVLD(Rs("BANKTRF_PMT"), 0)
      FEE_AMOUNT = NVLD(Rs("FEE_AMOUNT"), 0)
      
      OLD_EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
      DOCUMENT_SUB_TYPE_NAME = NVLS(Rs("DOCUMENT_SUB_TYPE_NAME"), "")
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      
      SALE_BY = NVLI(Rs("SALE_BY"), -1)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      
      CUS_PO = NVLS(Rs("CUS_PO"), "")
      DO_ADDRESS = NVLS(Rs("DO_ADDRESS"), "")
      
      BILLING_DOC_PACK = NVLI(Rs("BILLING_DOC_PACK"), -1)
      
   ElseIf Ind = 3 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 4 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      CREDIT = NVLD(Rs("CREDIT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
      NOTE = NVLS(Rs("NOTE"), "")
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      ADD_BRANCH_NAME = NVLS(Rs("ADD_BRANCH_NAME"), "")
      
   ElseIf Ind = 5 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      SALE_BY = NVLI(Rs("SALE_BY"), -1)
   ElseIf Ind = 7 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      CREDIT = NVLD(Rs("CREDIT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
      NOTE = NVLS(Rs("NOTE"), "")
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
   ElseIf Ind = 9 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
   ElseIf Ind = 10 Then
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
   ElseIf Ind = 11 Then
      APAR_TYPE_CODE = NVLS(Rs("APAR_TYPE_CODE"), "")
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
   ElseIf Ind = 12 Then
      'DO ใบส่งของ    รับคืน
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 13 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      APAR_TYPE_CODE = NVLS(Rs("APAR_TYPE_CODE"), "")
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
   ElseIf Ind = 14 Then
      'DO ใบส่งของ    รับคืน
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
   ElseIf Ind = 15 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      SR_REF_DO_NO = NVLS(Rs("SR_REF_DO_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      
      ITEM_AMOUNT_DOC = NVLD(Rs("ITEM_AMOUNT_DOC"), 0)
      AVG_PRICE_DOC = NVLD(Rs("AVG_PRICE_DOC"), 0)
      DISCOUNT_AMOUNT_DOC = NVLD(Rs("DISCOUNT_AMOUNT_DOC"), 0)
      TOTAL_PRICE_DOC = NVLD(Rs("TOTAL_PRICE_DOC"), 0)
      EXT_DISCOUNT_AMOUNT_DOC = NVLD(Rs("EXT_DISCOUNT_AMOUNT_DOC"), 0)
      
   ElseIf Ind = 17 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")

      SR_REF_DO_NO = NVLS(Rs("SR_REF_DO_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      
      BRANCH_ADDRESS = NVLI(Rs("BRANCH_ADDRESS"), -1)
      
      ITEM_AMOUNT_DOC = NVLD(Rs("ITEM_AMOUNT_DOC"), 0)
      AVG_PRICE_DOC = NVLD(Rs("AVG_PRICE_DOC"), 0)
      DISCOUNT_AMOUNT_DOC = NVLD(Rs("DISCOUNT_AMOUNT_DOC"), 0)
      TOTAL_PRICE_DOC = NVLD(Rs("TOTAL_PRICE_DOC"), 0)
      EXT_DISCOUNT_AMOUNT_DOC = NVLD(Rs("EXT_DISCOUNT_AMOUNT_DOC"), 0)
   ElseIf Ind = 18 Or Ind = 19 Or Ind = 129 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      NOTE = NVLS(Rs("NOTE"), "")
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      PO_NO = NVLS(Rs("PO_NO"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 20 Or Ind = 145 Or Ind = 146 Or Ind = 162 Then
      'DO ใบส่งของ    รับคืน
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      If Ind = 162 Then
         SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      End If
      If Ind = 20 Or Ind = 162 Then
         PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
         PAID_DISCOUNT = NVLD(Rs("PAID_DISCOUNT"), 0)
      End If
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      If Ind = 145 Or Ind = 146 Then
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         SALE_NAME = NVLS(Rs("SALE_NAME"), "")
         
         
         
         LOCATION_CODE = NVLS(Rs("LOCATION_CODE"), "")
      End If
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      If Ind = 20 Or Ind = 162 Then
         DOC_ID_BILLS_NO = NVLS(Rs("DOC_ID_BILLS_NO"), "")
         DOC_ID_BILLS_DUE_DATE = InternalDateToDate(NVLS(Rs("DOC_ID_BILLS_DUE_DATE"), ""))
      End If
   ElseIf Ind = 21 Then
      'DO ใบส่งของ    รับคืน
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      
      'ใบเสร็จรับชำระ
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 22 Then
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
      
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      APAR_TYPE_CODE = NVLS(Rs("APAR_TYPE_CODE"), "")
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
   ElseIf Ind = 23 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      
      If SHOW_ALL_FLAG <> "Y" Then
         PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
         PAID_DISCOUNT = NVLD(Rs("PAID_DISCOUNT"), 0)
      End If
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      
      CUS_PO = NVLS(Rs("CUS_PO"), "")
   ElseIf Ind = 24 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 25 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      'AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
   ElseIf Ind = 26 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
   ElseIf Ind = 27 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 28 Then
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      CNDN_AMOUNT = NVLD(Rs("CNDN_AMOUNT"), 0)
      
      CASH_PMT = NVLD(Rs("CASH_PMT"), 0)
      CHEQUE_PMT = NVLD(Rs("CHEQUE_PMT"), 0)
      BANKTRF_PMT = NVLD(Rs("BANKTRF_PMT"), 0)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      DOC_ID_NO = NVLS(Rs("DOC_ID_NO"), "")
      DOC_ID_DATE = InternalDateToDate(NVLS(Rs("DOC_ID_DATE"), ""))
      DOC_ID_PAID = NVLD(Rs("DOC_ID_PAID"), 0)
      DOC_ID1_PAID = NVLD(Rs("DOC_ID1_PAID"), 0)
      
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      CHEQUE_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), ""))
      EFFECTIVE_DATE = InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), ""))
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      DOCUMENT_TYPE_RCP = NVLI(Rs("DOCUMENT_TYPE_RCP"), -1)
      
   ElseIf Ind = 29 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      
   ElseIf Ind = 30 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")

      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
   ElseIf Ind = 31 Then
      BILLING_DOC_ID = NVLD(Rs("BILLING_DOC_ID"), 0)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      
      PO_NO = NVLS(Rs("PO_NO"), "")
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
   ElseIf Ind = 32 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
      
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      LOCATION_CODE = NVLS(Rs("LOCATION_CODE"), "")
   ElseIf Ind = 33 Then
      DOC_ITEM_ID = NVLI(Rs("DOC_ITEM_ID"), -1)
   ElseIf Ind = 34 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
         
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 35 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 36 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 37 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      
      
   ElseIf Ind = 38 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      
   ElseIf Ind = 39 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 40 Then
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      
      
      
   ElseIf Ind = 41 Then
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      
      
      
   ElseIf Ind = 42 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      
   ElseIf Ind = 43 Then
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      
      
      
   ElseIf Ind = 44 Then
            
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      
      
      
   ElseIf Ind = 45 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      
   ElseIf Ind = 46 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      
   ElseIf Ind = 47 Then
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
   ElseIf Ind = 48 Then
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      
      
   ElseIf Ind = 49 Then
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      
      
   ElseIf Ind = 50 Or Ind = 51 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      
   
   ElseIf Ind = 52 Then
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      
      
   ElseIf Ind = 53 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      
      
   ElseIf Ind = 54 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
   ElseIf Ind = 55 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
   ElseIf Ind = 56 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      NOT_SHOW_RETURN = NVLS(Rs("NOT_SHOW_RETURN"), "")
      
      
      
      
   ElseIf Ind = 57 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      NOT_SHOW_RETURN = NVLS(Rs("NOT_SHOW_RETURN"), "")
      
      
      
   ElseIf Ind = 58 Or Ind = 86 Or Ind = 116 Or Ind = 137 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      If Ind = 86 Then
         YYYYMM = NVLS(Rs("YYYYMM"), "")
         
       ElseIf Ind = 116 Then
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         
      ElseIf Ind = 137 Then
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      End If
   ElseIf Ind = 59 Or Ind = 87 Or Ind = 117 Or Ind = 138 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      If Ind = 87 Then
         YYYYMM = NVLS(Rs("YYYYMM"), "")
         
       ElseIf Ind = 117 Then
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         
      ElseIf Ind = 138 Then
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      End If
   ElseIf Ind = 122 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
   ElseIf Ind = 123 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
   ElseIf Ind = 60 Or Ind = 61 Or Ind = 158 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
   ElseIf Ind = 62 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
   ElseIf Ind = 63 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
   ElseIf Ind = 64 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
   ElseIf Ind = 65 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
   ElseIf Ind = 66 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 67 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 69 Or Ind = 71 Or Ind = 73 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      SUBTRACT_AMOUNT = NVLD(Rs("SUBTRACT_AMOUNT"), 0)
      ADDITION_AMOUNT = NVLD(Rs("ADDITION_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      
      RCP_DIF = NVLD(Rs("RCP_DIF"), 0)
      CASH_PMT = NVLD(Rs("CASH_PMT"), 0)
      CHEQUE_PMT = NVLD(Rs("CHEQUE_PMT"), 0)
      BANKTRF_PMT = NVLD(Rs("BANKTRF_PMT"), 0)
      FEE_AMOUNT = NVLD(Rs("FEE_AMOUNT"), 0)
      
      DOC_ID_DATE = InternalDateToDate(NVLS(Rs("DOC_ID_DATE"), ""))
      DOC_ID_NO = NVLS(Rs("DOC_ID_NO"), "")
      DOCUMENT_TYPE_RCP = NVLI(Rs("DOCUMENT_TYPE_RCP"), -1)
      DOC_ID_PAID = NVLD(Rs("DOC_ID_PAID"), 0)
      CNDN_AMOUNT = NVLD(Rs("CNDN_AMOUNT"), 0)
       
       If Ind = 73 Then
         SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      End If
   ElseIf Ind = 74 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      
      CUS_PO = NVLS(Rs("CUS_PO"), "")
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      EXT_DISCOUNT_PERCENT = NVLD(Rs("EXT_DISCOUNT_PERCENT"), 0)
   ElseIf Ind = 75 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 76 Then
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 77 Or Ind = 78 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      If Ind = 78 Then
         APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      End If
   ElseIf Ind = 79 Or Ind = 83 Then
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 80 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
   ElseIf Ind = 84 Or Ind = 85 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      If Ind = 84 Then
         APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      End If
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 88 Or Ind = 89 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
   ElseIf Ind = 90 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)   'DO ใบส่งของ    รับคืน
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)   'ใบเสร็จรับชำระ
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      SUBTRACT_AMOUNT = NVLD(Rs("SUBTRACT_AMOUNT"), 0)
      ADDITION_AMOUNT = NVLD(Rs("ADDITION_AMOUNT"), 0)
      
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 91 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 118 Or Ind = 119 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      NOT_SHOW_RETURN = NVLS(Rs("NOT_SHOW_RETURN"), "")
      
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      
      
      
   ElseIf Ind = 120 Or Ind = 121 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      NOT_SHOW_RETURN = NVLS(Rs("NOT_SHOW_RETURN"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 124 Or Ind = 128 Then
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
   ElseIf Ind = 125 Or Ind = 139 Then
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 126 Or Ind = 127 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
   ElseIf Ind = 130 Or Ind = 144 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      NOTE = NVLS(Rs("NOTE"), "")
   ElseIf Ind = 131 Or Ind = 132 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
   ElseIf Ind = 133 Or Ind = 134 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      
      
   ElseIf Ind = 135 Or Ind = 136 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      APAR_MAS_GROUP_CODE = NVLS(Rs("APAR_MAS_GROUP_CODE"), "")
      APAR_MAS_GROUP_NAME = NVLS(Rs("APAR_MAS_GROUP_NAME"), "")
   ElseIf Ind = 140 Or Ind = 141 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 142 Or Ind = 143 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 163 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
   ElseIf Ind = 164 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      Due_Date = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
   End If
End Sub
Public Sub PopulateFromRS2(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 92 Then
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 93 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")

      SR_REF_DO_NO = NVLS(Rs("SR_REF_DO_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      
      ITEM_AMOUNT_DOC = NVLD(Rs("ITEM_AMOUNT_DOC"), 0)
      AVG_PRICE_DOC = NVLD(Rs("AVG_PRICE_DOC"), 0)
      DISCOUNT_AMOUNT_DOC = NVLD(Rs("DISCOUNT_AMOUNT_DOC"), 0)
      TOTAL_PRICE_DOC = NVLD(Rs("TOTAL_PRICE_DOC"), 0)
      EXT_DISCOUNT_AMOUNT_DOC = NVLD(Rs("EXT_DISCOUNT_AMOUNT_DOC"), 0)
   ElseIf Ind = 94 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_TYPE_SUB_CODE = NVLS(Rs("STOCK_TYPE_SUB_CODE"), "")
      STOCK_TYPE_SUB_NAME = NVLS(Rs("STOCK_TYPE_SUB_NAME"), "")
      
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 95 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      CUS_PO = NVLS(Rs("CUS_PO"), "")
   ElseIf Ind = 96 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 97 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      
   ElseIf Ind = 98 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      
      
   ElseIf Ind = 99 Or Ind = 114 Or Ind = 115 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
   ElseIf Ind = 100 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_SUB_TYPE = NVLI(Rs("DOCUMENT_SUB_TYPE"), -1)
   ElseIf Ind = 101 Then
      PO_ID = NVLI(Rs("PO_ID"), -1)
   ElseIf Ind = 102 Or Ind = 103 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
   ElseIf Ind = 104 Or Ind = 105 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
   ElseIf Ind = 110 Or Ind = 111 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      
      
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 147 Or Ind = 148 Or Ind = 159 Then
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      If Ind = 147 Or Ind = 148 Then
         NOT_SHOW_RETURN = NVLS(Rs("NOT_SHOW_RETURN"), "")
      End If
   ElseIf Ind = 149 Or Ind = 150 Or Ind = 160 Or Ind = 161 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 151 Or Ind = 152 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 153 Or Ind = 154 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 155 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 156 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
   ElseIf Ind = 157 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
   ElseIf Ind = 165 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
 ElseIf Ind = 166 Then
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
   ElseIf Ind = 167 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      APAR_TYPE = NVLI(Rs("APAR_TYPE"), -1)
      
      STOCK_GROUP = NVLI(Rs("STOCK_GROUP"), -1)
   ElseIf Ind = 168 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
   End If
End Sub
Public Function UpDatePaidAmount() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "    PAID_AMOUNT = " & Val(PAID_AMOUNT) & "," & _
                  "    PAY_AMOUNT = " & Val(PAY_AMOUNT) & ", " & _
                  "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & Val(BILLING_DOC_ID)
   
   Call m_ADOConn.Execute(SQL1)

End Function
Public Function UpdateDocSubType() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
Dim TempID13 As String
   TempID13 = DOCUMENT_SUB_TYPE
   If DOCUMENT_SUB_TYPE <= 0 Then
      TempID13 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   DOCUMENT_SUB_TYPE = " & TempID13 & "," & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & Val(BILLING_DOC_ID)
   
   Call m_ADOConn.Execute(SQL1)
   
End Function
Public Sub DeleteDataFromBillingPack()
Dim SQL As String
Dim Tf As CTableField
Dim iCount As Long
Dim TempRs As ADODB.Recordset
   
   SQL = "DELETE FROM CASH_TRAN WHERE BILLING_DOC_PACK = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM CHEQUE WHERE BILLING_DOC_PACK = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM RCPCNDN_ITEM WHERE BILLING_DOC_PACK = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_SUBTRACT WHERE BILLING_DOC_PACK = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_ADDITION WHERE BILLING_DOC_PACK = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_PACK = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "UPDATE BILLING_DOC SET REFER_DESC = NULL WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub AddSearchLocationNo(FromStr As String)
   If InStr(1, FromStr, "LC") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF LC ON (DI.LOCATION_ID = LC.KEY_ID) "
   End If
End Sub
Public Sub AddSearchStockCode(FromStr As String)
   If InStr(1, FromStr, "SC") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4  As String
Dim TempID6  As String
Dim TempID8  As String
Dim TempID9  As String
Dim TempID10  As String
Dim TempID11  As String
Dim TempID12  As String
Dim TempID13 As String

   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   TempID1 = BILLING_ADDRESS_ID
   If BILLING_ADDRESS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = ENTERPRISE_ADDRESS_ID
   If ENTERPRISE_ADDRESS_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = APAR_MAS_ID
   If APAR_MAS_ID <= 0 Then
      TempID4 = "NULL"
   End If
      
   TempID6 = SALE_BY
   If SALE_BY <= 0 Then
      TempID6 = "NULL"
   End If
   
   TempID9 = BRANCH_ADDRESS
   If BRANCH_ADDRESS <= 0 Then
      TempID9 = "NULL"
   End If
   
   TempID10 = SR_REF_DO_ID
   If SR_REF_DO_ID <= 0 Then
      TempID10 = "NULL"
   End If
   
   TempID11 = BILLING_DOC_PACK
   If BILLING_DOC_PACK <= 0 Then
      TempID11 = "NULL"
   End If
   
   TempID13 = DOCUMENT_SUB_TYPE
   If DOCUMENT_SUB_TYPE <= 0 Then
      TempID13 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILLING_DOC_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BILLING_DOC " & _
                     "( " & _
                     "   BILLING_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   DUE_DATE, " & _
                     "   NOTE, " & _
                     "   BILLING_ADDRESS_ID, " & _
                     "   ENTERPRISE_ADDRESS_ID, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   APAR_MAS_ID, " & _
                     "   TOTAL_AMOUNT, " & _
                     "   TOTAL_PRICE, " & _
                     "   VAT_PERCENT, " & _
                     "   VAT_AMOUNT, " & _
                     "   WH_PERCENT, " & _
                     "   WH_AMOUNT, "
            SQL1 = SQL1 & _
                     "   DISCOUNT_AMOUNT, " & _
                     "   CREDIT, " & _
                     "   DEBIT_AMOUNT, " & _
                     "   CREDIT_AMOUNT, " & _
                     "   SUBTRACT_AMOUNT, " & _
                     "   ADDITION_AMOUNT, " & _
                     "   EXT_DISCOUNT_AMOUNT, " & _
                     "   EXT_DISCOUNT_PERCENT, " & _
                     "   PAY_AMOUNT, " & _
                     "   PAID_AMOUNT, " & _
                     "   SALE_BY, " & _
                     "   CUS_PO, " & _
                     "   DO_ADDRESS, " & _
                     "   BRANCH_ADDRESS, " & _
                     "   DOCUMENT_SUB_TYPE, "
               SQL1 = SQL1 & _
                     "   CANCEL_FLAG, " & _
                     "   SR_REF_DO_ID, " & _
                     "   SR_REF_DO_NO, " & _
                     "   YYYYMM, " & _
                     "   RCP_CASH_TRAN, " & _
                     "   RCP_DIF, " & _
                     "   BILLING_DOC_PACK, " & _
                     "   CASH_PMT, " & _
                     "   CHEQUE_PMT, " & _
                     "   BANKTRF_PMT, " & _
                     "   FEE_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                        DOCUMENT_TYPE & "," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(Due_Date))) & "'," & _
                        "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                        TempID1 & ","
         SQL1 = SQL1 & _
                        TempID2 & "," & _
                        TempID3 & "," & _
                        TempID4 & "," & _
                        TOTAL_AMOUNT & "," & _
                        TOTAL_PRICE & "," & _
                        VAT_PERCENT & "," & _
                        VAT_AMOUNT & "," & _
                        WH_PERCENT & "," & _
                        WH_AMOUNT & "," & _
                        DISCOUNT_AMOUNT & "," & _
                        CREDIT & "," & _
                        DEBIT_AMOUNT & "," & _
                        CREDIT_AMOUNT & "," & _
                        SUBTRACT_AMOUNT & "," & _
                        ADDITION_AMOUNT & "," & _
                        EXT_DISCOUNT_AMOUNT & "," & _
                        EXT_DISCOUNT_PERCENT & ","
                  SQL1 = SQL1 & _
                        PAY_AMOUNT & "," & _
                        PAID_AMOUNT & "," & _
                        TempID6 & "," & _
                        "'" & ChangeQuote(Trim(CUS_PO)) & "'," & _
                        "'" & ChangeQuote(Trim(DO_ADDRESS)) & "'," & _
                        TempID9 & "," & _
                        TempID13 & "," & _
                        "'" & ChangeQuote(Trim(CANCEL_FLAG)) & "'," & _
                        TempID10 & "," & _
                        "'" & ChangeQuote(Trim(SR_REF_DO_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        RCP_CASH_TRAN & "," & _
                        RCP_DIF & "," & _
                        TempID11 & "," & _
                        CASH_PMT & ","
                  SQL1 = SQL1 & _
                        CHEQUE_PMT & "," & _
                        BANKTRF_PMT & "," & _
                        FEE_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BILLING_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BILLING_DOC SET " & _
                     "    DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   DOCUMENT_DATE= '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "    DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "   DUE_DATE= '" & ChangeQuote(Trim(DateToStringInt(Due_Date))) & "', " & _
                     "    NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "    BILLING_ADDRESS_ID = " & TempID1 & ", " & _
                     "    ENTERPRISE_ADDRESS_ID = " & TempID2 & ", " & _
                     "    APAR_MAS_ID = " & TempID4 & ", " & _
                     "    TOTAL_AMOUNT = " & TOTAL_AMOUNT & ", " & _
                     "    TOTAL_PRICE = " & TOTAL_PRICE & "," & _
                     "    VAT_PERCENT = " & VAT_PERCENT & "," & _
                     "    VAT_AMOUNT = " & VAT_AMOUNT & "," & _
                     "    WH_PERCENT = " & WH_PERCENT & ", " & _
                     "    WH_AMOUNT = " & WH_AMOUNT & ", " & _
                     "    DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & "," & _
                     "    CREDIT = " & CREDIT & ", " & _
                     "    DEBIT_AMOUNT = " & DEBIT_AMOUNT & ", " & _
                     "    CREDIT_AMOUNT = " & CREDIT_AMOUNT & ","
            SQL1 = SQL1 & _
                     "    SUBTRACT_AMOUNT = " & SUBTRACT_AMOUNT & ", " & _
                     "    ADDITION_AMOUNT = " & ADDITION_AMOUNT & ", " & _
                     "    EXT_DISCOUNT_AMOUNT = " & EXT_DISCOUNT_AMOUNT & ", " & _
                     "    EXT_DISCOUNT_PERCENT = " & EXT_DISCOUNT_PERCENT & ", " & _
                     "    PAY_AMOUNT = " & PAY_AMOUNT & ", " & _
                     "    PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "    SALE_BY = " & TempID6 & ", " & _
                     "    CUS_PO = '" & ChangeQuote(Trim(CUS_PO)) & "', " & _
                     "    DO_ADDRESS = '" & ChangeQuote(Trim(DO_ADDRESS)) & "', " & _
                     "    BRANCH_ADDRESS = " & TempID9 & ", " & _
                     "    DOCUMENT_SUB_TYPE = " & TempID13 & ", " & _
                     "    CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "', " & _
                     "    SR_REF_DO_ID = " & TempID10 & ", " & _
                     "    SR_REF_DO_NO = '" & ChangeQuote(Trim(SR_REF_DO_NO)) & "', " & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    RCP_CASH_TRAN = " & RCP_CASH_TRAN & ", "
               SQL1 = SQL1 & _
                     "    RCP_DIF = " & RCP_DIF & ", " & _
                     "    CASH_PMT = " & CASH_PMT & ", " & _
                     "    CHEQUE_PMT = " & CHEQUE_PMT & ", " & _
                     "    BANKTRF_PMT = " & BANKTRF_PMT & ", " & _
                     "    FEE_AMOUNT = " & FEE_AMOUNT & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function


