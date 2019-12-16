VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrintLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String

Public DOCUMENT_TYPE  As Long
Public EMP_ID  As Long
Public PART_ITEM_ID As Long
Public BRANCH_ID As Long
Public TOTAL_AMOUNT   As Double
Public TOTAL_PRICE As Double
Public DISCOUNT_AMOUNT As Double
Public EXTRA_DISCOUNT_AMOUNT As Double

Public PART_DESC As String
Public FREE_FLAG As String
Public CANCEL_FLAG As String

Public LABEL_DESC As String
Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public Due_Date As Date
Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String
Public FROM_SALE_CODE  As String
Public TO_SALE_CODE  As String
Public FROM_APAR_CODE As String
Public TO_APAR_CODE As String

Public SALE_ID  As Long

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
      
   CANCEL_FLAG = "N"
   
   Call RegisterFields
End Sub
Private Sub Class_Terminate()
   Set m_FieldList = Nothing
End Sub
Private Sub AddField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FName, FType, FValue, FCat)
   Call m_FieldList.add(F, FName)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldName As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   GetFieldValue = Tf.GetValue
End Function

Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.SetValue (Value)
End Sub

Public Sub AddEditData()
Dim SQL As String

   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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

Dim PRINT_LABEL_ID As Long
Dim DOC_ITEM_ID As Long
Dim BILLING_DOC_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim BRANCH_ID As Long

Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim STOCK_NO As String

   PRINT_LABEL_ID = Val(GetFieldValue("PRINT_LABEL_ID"))
   DOC_ITEM_ID = Val(GetFieldValue("DOC_ITEM_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   STOCK_NO = GetFieldValue("STOCK_NO")
   BRANCH_ID = Val(GetFieldValue("BRANCH_ID"))
   EMP_ID = Val(GetFieldValue("EMP_ID"))              ' SALE_ID
   SALE_ID = Val(GetFieldValue("EMP_ID"))

   If Ind = 1 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME BLOCK_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_CODE BRANCH_CODE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME BRANCH_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_ID,EM.EMP_CODE SALE_CODE,NM.LONG_NAME SALE_LONG_NAME,NM.LAST_NAME SALE_LAST_NAME "
       SelectStr1 = SelectStr1 & ",MR3.KEY_NAME  UNIT_CHANGE_NAME "                       '''''''''''''''''''''''''PAT
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.BLOCK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EN.SALE_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EM.EMP_ID = EN2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN2.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (EN.UNIT_TRAN_ID = MR3.KEY_ID) "              '''''''''''''''''''''''''PAT
      
      CANCEL_FLAG = ""
      WhereStr = ""
   ElseIf Ind = 2 Then     ' CReportPrintLabel001  CReportPrintLabel004
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",SC.BILL_DESC LABEL_DESC,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME BLOCK_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME BRANCH_NAME,MR2.KEY_CODE BRANCH_CODE "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "      '''''''''''''''''''''''''PAT
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.BLOCK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
     FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (EN.UNIT_TRAN_ID = MR3.KEY_ID) " '''''''''''''''''''''''''PAT
      WhereStr = ""
      CANCEL_FLAG = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT EN.BLOCK_ID,EN.BRANCH_ID "
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME BLOCK_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME BRANCH_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.BLOCK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      
      WhereStr = ""
      CANCEL_FLAG = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = ""
      CANCEL_FLAG = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_CODE_ID PART_ITEM_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = ""
      CANCEL_FLAG = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DISTINCT EN.BLOCK_ID,DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME BLOCK_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (EN.BLOCK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      
      WhereStr = ""
      CANCEL_FLAG = ""
   ElseIf Ind = 7 Then 'Creport BillingDoc003
      SelectStr1 = "SELECT SUM(EN.TOTAL_PRICE_SUM) TOTAL_PRICE,SUM(EN.TOTAL_AMOUNT_SUM) TOTAL_AMOUNT,MR2.KEY_NAME BRANCH_NAME "
      SelectStr1 = SelectStr1 & ",SUM(EN.DISCOUNT_AMOUNT_SUM) DISCOUNT_AMOUNT,SUM(EN.EXTRA_DISCOUNT_AMOUNT_SUM) EXTRA_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",MR6.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EN.SALE_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (SC.UNIT_CHANGE_ID = MR6.KEY_ID) "
      
      WhereStr = "WHERE BD.DOCUMENT_TYPE IN ( " & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ") "
   ElseIf Ind = 8 Then 'Creport BillingDoc004
      SelectStr1 = "SELECT SUM(EN.TOTAL_PRICE_SUM) TOTAL_PRICE,SUM(EN.TOTAL_AMOUNT_SUM) TOTAL_AMOUNT,MR2.KEY_NAME BRANCH_NAME "
      SelectStr1 = SelectStr1 & ",SUM(EN.DISCOUNT_AMOUNT_SUM) DISCOUNT_AMOUNT,SUM(EN.EXTRA_DISCOUNT_AMOUNT_SUM) EXTRA_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",DI.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EN.SALE_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
         
      WhereStr = "WHERE BD.DOCUMENT_TYPE IN ( " & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ") "
   ElseIf Ind = 9 Then 'Creport BillingDoc005
      SelectStr1 = "SELECT DISTINCT MR2.KEY_ID BRANCH_ID,MR2.KEY_CODE BRANCH_CODE,MR2.KEY_NAME BRANCH_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_ID,EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR2.PARENT_EX_ID2 = APM.APAR_MAS_ID ) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EN.SALE_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
         
      WhereStr = "WHERE BD.DOCUMENT_TYPE IN ( " & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ") "
   
   ElseIf Ind = 12 Then 'Creport Commission001
      SelectStr1 = "SELECT SUM(EN.TOTAL_PRICE) TOTAL_PRICE,SUM(EN.TOTAL_AMOUNT) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(EN.DISCOUNT_AMOUNT_SUM) DISCOUNT_AMOUNT,SUM(EN.EXTRA_DISCOUNT_AMOUNT_SUM) EXTRA_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_ID,SC.GROUP_COM_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (EN.BRANCH_ID = MR2.KEY_ID) "
      'FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EN.SALE_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (EN.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
         
      WhereStr = "WHERE BD.DOCUMENT_TYPE IN ( " & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ") "
   ElseIf Ind = 13 Then  'Patch_2007_12_01_1_jill
      SelectStr1 = "SELECT EN.* "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      WhereStr = ""
   End If
   
   If PRINT_LABEL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.PRINT_LABEL_ID = " & PRINT_LABEL_ID & ") " & vbCrLf
   End If
   
   If DOC_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.DOC_ITEM_ID = " & DOC_ITEM_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (DI.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If BRANCH_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BRANCH_ID = " & BRANCH_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(FROM_STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
   End If
   
   If Len(Trim(TO_STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
   End If

   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
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
   
   If Len(Trim(FREE_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(FREE_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(DI.FREE_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = '" & ChangeQuote(Trim(FREE_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CANCEL_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(BD.CANCEL_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (BD.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   If Ind = 7 Then
      GroupByStr = "GROUP BY MR2.KEY_NAME,MR6.KEY_NAME,BD.DOCUMENT_TYPE "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY MR2.KEY_NAME,DI.PART_ITEM_ID "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY EM.EMP_ID,MR2.KEY_ID,DI.PART_ITEM_ID,BD.DOCUMENT_TYPE "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY EM.EMP_ID,SC.GROUP_COM_ID "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY PRINT_LABEL_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE,MR2.KEY_CODE,SC.REPORT_PRIORITY "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY EN.BLOCK_ID,EN.BRANCH_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY SC.STOCK_CODE_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PRINT_LABEL_ID "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY EN.BLOCK_ID,DI.PART_ITEM_ID,EN.BRANCH_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY MR2.KEY_NAME "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY DI.PART_ITEM_ID,MR2.KEY_NAME "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY EM.EMP_ID,MR2.KEY_CODE "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY EM.EMP_ID "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY PRINT_LABEL_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub QueryDataReport(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
   
   If Ind = 10 Then 'LoadDistinctLabelPartItemEx
      SelectStr1 = "SELECT DISTINCT DI.PART_ITEM_ID,SC.BILL_DESC LABEL_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " PL "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (PL.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
         
      WhereStr = "WHERE BD.DOCUMENT_TYPE IN ( " & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ") "
   ElseIf Ind = 11 Then 'LoadSumBranchPartItem
      SelectStr1 = "SELECT SUM(PL.TOTAL_PRICE_SUM) TOTAL_PRICE,SUM(PL.TOTAL_AMOUNT_SUM) TOTAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(PL.DISCOUNT_AMOUNT_SUM) DISCOUNT_AMOUNT,SUM(PL.EXTRA_DISCOUNT_AMOUNT_SUM) EXTRA_DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",EM.EMP_ID "
      SelectStr1 = SelectStr1 & ",DI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",MR2.KEY_ID BRANCH_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " PL "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (PL.BRANCH_ID = MR2.KEY_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MR2.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (PL.SALE_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOC_ITEM DI ON (PL.DOC_ITEM_ID = DI.DOC_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (DI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = "WHERE BD.DOCUMENT_TYPE IN ( " & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ") "
   End If
   
   If BRANCH_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PL.BRANCH_ID = " & BRANCH_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(FROM_STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
   End If
   
   If Len(Trim(TO_STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
   End If

   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
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
   
   If Len(Trim(FREE_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(FREE_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(DI.FREE_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (DI.FREE_FLAG = '" & ChangeQuote(Trim(FREE_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CANCEL_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(BD.CANCEL_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (BD.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   If Ind = 11 Then
      GroupByStr = "GROUP BY EM.EMP_ID,MR2.KEY_ID,DI.PART_ITEM_ID "
   End If
   
   If Ind = 10 Then
      OrderByStr = "ORDER BY SC.BILL_DESC "
   ElseIf Ind = 2 Then
      'OrderByStr = "ORDER BY MR1.KEY_CODE,MR2.KEY_CODE,SC.REPORT_PRIORITY "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   If Ind = 11 Then
      OrderByStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("PRINT_LABEL_ID", NVLI(Rs("PRINT_LABEL_ID"), -1))
      Call SetFieldValue("DOC_ITEM_ID", NVLI(Rs("DOC_ITEM_ID"), -1))
      Call SetFieldValue("BLOCK_ID", NVLI(Rs("BLOCK_ID"), -1))
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("PACK_AMOUNT", NVLD(Rs("PACK_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      
      Call SetFieldValue("UNIT_TRAN_ID", NVLI(Rs("UNIT_TRAN_ID"), -1))
      Call SetFieldValue("UNIT_MULTIPLE", NVLD(Rs("UNIT_MULTIPLE"), 0))
      
      Call SetFieldValue("BLOCK_NAME", NVLS(Rs("BLOCK_NAME"), ""))
      Call SetFieldValue("BRANCH_CODE", NVLS(Rs("BRANCH_CODE"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      
      Call SetFieldValue("SALE_ID", NVLI(Rs("SALE_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("SALE_CODE", NVLS(Rs("SALE_CODE"), ""))
      Call SetFieldValue("SALE_LONG_NAME", NVLS(Rs("SALE_LONG_NAME"), ""))
      Call SetFieldValue("SALE_LAST_NAME", NVLS(Rs("SALE_LAST_NAME"), ""))
      
   ElseIf Ind = 2 Then
      Call SetFieldValue("PRINT_LABEL_ID", NVLI(Rs("PRINT_LABEL_ID"), -1))
      Call SetFieldValue("DOC_ITEM_ID", NVLI(Rs("DOC_ITEM_ID"), -1))
      Call SetFieldValue("BLOCK_ID", NVLI(Rs("BLOCK_ID"), -1))
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("PACK_AMOUNT", NVLD(Rs("PACK_AMOUNT"), 0))
      
      Call SetFieldValue("BLOCK_NAME", NVLS(Rs("BLOCK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("BRANCH_CODE", NVLS(Rs("BRANCH_CODE"), ""))

      Call SetFieldValue("LABEL_DESC", NVLS(Rs("LABEL_DESC"), ""))
      Call SetFieldValue("PART_DESC", NVLS(Rs("PART_DESC"), ""))
      
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))                '''''''''''''''''''''''''PAT
      Call SetFieldValue("UNIT_MULTIPLE", NVLD(Rs("UNIT_MULTIPLE"), 0))                '''''''''''''''''''''''''PAT
      Call SetFieldValue("UNIT_CHANGE_NAME", NVLS(Rs("UNIT_CHANGE_NAME"), ""))                '''''''''''''''''''''''''PAT
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))                '''''''''''''''''''''''''PAT
   ElseIf Ind = 3 Then
      Call SetFieldValue("BLOCK_ID", NVLI(Rs("BLOCK_ID"), -1))
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      
      Call SetFieldValue("BLOCK_NAME", NVLS(Rs("BLOCK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
   ElseIf Ind = 4 Then
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("PART_NO", NVLS(Rs("PART_NO"), ""))
      Call SetFieldValue("PART_DESC", NVLS(Rs("PART_DESC"), ""))
   ElseIf Ind = 5 Then
      Call SetFieldValue("PRINT_LABEL_ID", NVLI(Rs("PRINT_LABEL_ID"), -1))
      Call SetFieldValue("DOC_ITEM_ID", NVLI(Rs("DOC_ITEM_ID"), -1))
      Call SetFieldValue("BLOCK_ID", NVLI(Rs("BLOCK_ID"), -1))
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("PACK_AMOUNT", NVLD(Rs("PACK_AMOUNT"), 0))
      
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
   ElseIf Ind = 6 Then
      Call SetFieldValue("BLOCK_ID", NVLI(Rs("BLOCK_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      
      Call SetFieldValue("BLOCK_NAME", NVLS(Rs("BLOCK_NAME"), ""))
   ElseIf Ind = 7 Then
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("UNIT_CHANGE_NAME", NVLS(Rs("UNIT_CHANGE_NAME"), ""))
   ElseIf Ind = 8 Then
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("DISCOUNT_AMOUNT", NVLD(Rs("DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("EXTRA_DISCOUNT_AMOUNT", NVLD(Rs("EXTRA_DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
   ElseIf Ind = 9 Then
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      Call SetFieldValue("BRANCH_CODE", NVLS(Rs("BRANCH_CODE"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      
      Call SetFieldValue("SALE_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("SALE_CODE", NVLS(Rs("SALE_CODE"), ""))
      Call SetFieldValue("SALE_LONG_NAME", NVLS(Rs("SALE_LONG_NAME"), ""))
      Call SetFieldValue("SALE_LAST_NAME", NVLS(Rs("SALE_LAST_NAME"), ""))
      
   ElseIf Ind = 10 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LABEL_DESC = NVLS(Rs("LABEL_DESC"), "")
   ElseIf Ind = 11 Then
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXTRA_DISCOUNT_AMOUNT = NVLD(Rs("EXTRA_DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      Call SetFieldValue("SALE_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("GROUP_COM_ID", NVLI(Rs("GROUP_COM_ID"), -1))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("DISCOUNT_AMOUNT", NVLD(Rs("DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("EXTRA_DISCOUNT_AMOUNT", NVLD(Rs("EXTRA_DISCOUNT_AMOUNT"), 0))
      
   ElseIf Ind = 13 Then
      Call SetFieldValue("PRINT_LABEL_ID", NVLI(Rs("PRINT_LABEL_ID"), -1))
      Call SetFieldValue("DOC_ITEM_ID", NVLI(Rs("DOC_ITEM_ID"), -1))
      Call SetFieldValue("BLOCK_ID", NVLI(Rs("BLOCK_ID"), -1))
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("PACK_AMOUNT", NVLD(Rs("PACK_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      
      Call SetFieldValue("DISCOUNT_AMOUNT", NVLD(Rs("DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("EXTRA_DISCOUNT_AMOUNT", NVLD(Rs("EXTRA_DISCOUNT_AMOUNT"), 0))
      
      Call SetFieldValue("UNIT_TRAN_ID", NVLI(Rs("UNIT_TRAN_ID"), -1))
      Call SetFieldValue("UNIT_MULTIPLE", NVLD(Rs("UNIT_MULTIPLE"), 0))
      
   End If
End Sub

Public Sub RegisterFields()
   TableName = "PRINT_LABEL"
   SequenceName = "PRINT_LABEL_SEQ"
   PKField = "PRINT_LABEL_ID"
   
   If MasterInd = "1" Then
      Call AddField("PRINT_LABEL_ID", INT_TYPE, -1, ID_CAT)
      Call AddField("DOC_ITEM_ID", INT_TYPE, "", DATA_CAT)
      Call AddField("BLOCK_ID", INT_TYPE, "", DATA_CAT)
      Call AddField("BRANCH_ID", INT_TYPE, "", DATA_CAT)
      Call AddField("SALE_ID", INT_TYPE, "", DATA_CAT)
      
      Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("PACK_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TOTAL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("DISCOUNT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("EXTRA_DISCOUNT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      
      Call AddField("TOTAL_PRICE_SUM", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TOTAL_AMOUNT_SUM", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("DISCOUNT_AMOUNT_SUM", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("EXTRA_DISCOUNT_AMOUNT_SUM", MONEY_TYPE, 0, DATA_CAT)
      
      Call AddField("UNIT_TRAN_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("UNIT_MULTIPLE", MONEY_TYPE, 0, DATA_CAT)
         
      Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
      Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
      Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
      Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
      
      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
      
       Call AddField("BLOCK_NAME", STRING_TYPE, "", TEMP_CAT)
       Call AddField("BRANCH_CODE", STRING_TYPE, "", TEMP_CAT)
       Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
       
       Call AddField("BILLING_DOC_ID", INT_TYPE, "", TEMP_CAT)
       Call AddField("LABEL_DESC", STRING_TYPE, "", TEMP_CAT)
       
       Call AddField("PART_ITEM_ID", INT_TYPE, "", TEMP_CAT)
       Call AddField("PART_NO", STRING_TYPE, "", TEMP_CAT)
       Call AddField("PART_DESC", STRING_TYPE, "", TEMP_CAT)
       
      Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
      
      Call AddField("STOCK_NO", STRING_TYPE, "", TEMP_CAT)
       
       Call AddField("LOCATION_SALE", INT_TYPE, "", TEMP_CAT)
       Call AddField("LOCATION_SALE_CODE", STRING_TYPE, "", TEMP_CAT)
       Call AddField("LOCATION_SALE_NAME", STRING_TYPE, "", TEMP_CAT)
       
      Call AddField("EMP_ID", INT_TYPE, "", TEMP_CAT)
      Call AddField("SALE_CODE", STRING_TYPE, "", TEMP_CAT)
      Call AddField("SALE_LONG_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("SALE_LAST_NAME", STRING_TYPE, "", TEMP_CAT)
   
      Call AddField("UNIT_TRAN_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("UNIT_CHANGE_NAME", STRING_TYPE, "", TEMP_CAT)
      
      Call AddField("GROUP_COM_ID", INT_TYPE, -1, TEMP_CAT)
   End If
End Sub

