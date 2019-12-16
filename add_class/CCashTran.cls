VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCashTran"
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
Public QueryFlag As Long

Public CASH_TRAN_ID As Long
Public BILLING_DOC_ID As Long
Public CASH_DOC_ID As Long

Public TX_TYPE As String
Public BANK_ACCOUNT As Long
Public BANK_FLAG As String
Public POST_FLAG As String
Public CHEQUE_ID As Long
Public BANK_ID As Long
Public BANK_BRANCH As Long
Public NET_AMOUNT As Double
'---------------------------------------------------------------------------------------------------------------------------->

Public FROM_DATE As Date
Public TO_DATE As Date
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public APAR_CODE As String
Public BANK_NAME As String
Public BRANCH_NAME As String
Public CHEQUE_NO As String
Public CHEQUE_DATE  As Date
Public CHEQUE_AMOUNT  As Double

Public BANK_ACCOUNT_NAME As String
Public Cheque As CCheque
Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set Cheque = New CCheque
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set Cheque = Nothing
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
Dim Tf As CTableField
Dim NET_AMOUNT As Double
Dim FEE_AMOUNT As Double
Dim Amount As Double

   Set Tf = m_FieldList(PKField)
   
   If GetFieldValue("FROM_BILLING") = "" Then
      Call SetFieldValue("FROM_BILLING", "N")
   End If
   
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

Dim CASH_TRAN_ID As Long
Dim BILLING_DOC_ID As Long
Dim CASH_DOC_ID As Long

Dim TX_TYPE As String
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim APAR_CODE As String
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim BANK_ACCOUNT As Long
Dim BANK_FLAG As String
Dim POST_FLAG As String

   CASH_TRAN_ID = Val(GetFieldValue("CASH_TRAN_ID"))
   CASH_DOC_ID = Val(GetFieldValue("CASH_DOC_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   TX_TYPE = GetFieldValue("TX_TYPE")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   APAR_CODE = GetFieldValue("APAR_CODE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   BANK_ACCOUNT = Val(GetFieldValue("BANK_ACCOUNT"))
   BANK_FLAG = GetFieldValue("BANK_FLAG")
   POST_FLAG = GetFieldValue("POST_FLAG")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.KEY_NAME BANK_NAME, MR2.KEY_NAME BRANCH_NAME, MR3.KEY_NAME ACCOUNT_NAME, CQ.CHEQUE_NO, CQ.CHEQUE_DATE, CQ.CHEQUE_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.BANK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.BANK_BRANCH = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.BANK_ACCOUNT = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE CQ ON (NM.CHEQUE_ID = CQ.CHEQUE_ID) "
      
      WhereStr = ""
   
   End If
   
   If CASH_TRAN_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_TRAN_ID = " & CASH_TRAN_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
      
   If CASH_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_ID = " & CASH_DOC_ID & ") " & vbCrLf
   End If
      
   If BANK_ACCOUNT > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BANK_ACCOUNT = " & BANK_ACCOUNT & ") " & vbCrLf
   End If
            
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.TX_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.TX_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If Trim(APAR_CODE) <> "" Then
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(BANK_FLAG) <> "" Then
      If WildCard(BANK_FLAG, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CQ.BANK_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CQ.BANK_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CQ.BANK_FLAG = '" & ChangeQuote(Trim(BANK_FLAG)) & "')"
         Else
            WhereStr = WhereStr & "AND (CQ.BANK_FLAG = '" & ChangeQuote(Trim(BANK_FLAG)) & "')"
         End If
      End If
   End If
   
   If Trim(POST_FLAG) <> "" Then
      If WildCard(POST_FLAG, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CQ.POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CQ.POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CQ.POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
         Else
            WhereStr = WhereStr & "AND (CQ.POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
         End If
      End If
   End If
   
   If Trim(TX_TYPE) <> "" Then
      If WildCard(TX_TYPE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.TX_TYPE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.TX_TYPE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "')"
         End If
      End If
   End If
   
   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = "GROUP BY NM.BANK_ACCOUNT, NM.TX_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY NM.TX_DATE, NM.APAR_MAS_ID, NM.BANK_ACCOUNT "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY NM.APAR_MAS_ID, NM.PAYMENT_TYPE, NM.TX_DATE "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY NM.APAR_MAS_ID, NM.PAYMENT_TYPE, NM.TX_DATE "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY NM.BILLING_DOC_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY BD.RECEIPT_TYPE,NM.PAYMENT_TYPE "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY NM.PAYMENT_TYPE "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY NM.TX_NO,NM.TX_DATE,MR1.KEY_NAME , MR2.KEY_NAME , MR3.KEY_NAME "
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY CASH_TRAN_ID "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY CASH_TRAN_ID "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY MR3.KEY_CODE ASC, MR3.KEY_NAME ASC, NM.TX_DATE ASC, NM.TX_NO "
      Else
         OrderByStr = "ORDER BY CASH_TRAN_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY MR3.KEY_CODE ASC, MR3.KEY_CODE "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY NM.BANK_ACCOUNT "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY NM.TX_DATE ASC, CM.APAR_CODE ASC, NM1.LONG_NAME "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY NM.TX_DATE "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY CQ.CHEQUE_DATE "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY NM.APAR_MAS_ID "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY NM.APAR_MAS_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY NM.BILLING_DOC_ID "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY NM.PAYMENT_TYPE "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY NM.PAYMENT_TYPE "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY NM.TX_NO,NM.TX_DATE "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY MR3.KEY_NAME "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub QueryDataReport(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional CountFlag As Boolean = False)
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
   
   If Ind = 2 Then   'frmAddChequeItemEx
      SelectStr1 = "SELECT MR1.KEY_NAME BANK_NAME, MR2.KEY_NAME BRANCH_NAME, CQ.CHEQUE_NO, CQ.CHEQUE_DATE, CQ.CHEQUE_AMOUNT "
      SelectStr1 = SelectStr1 & ",CQ.BANK_ID,CQ.BANK_BRANCH,NM.CHEQUE_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE CQ ON (NM.CHEQUE_ID = CQ.CHEQUE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (CQ.BANK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (CQ.BANK_BRANCH = MR2.KEY_ID) "
      
      WhereStr = " WHERE  (NM.CASH_DOC_ID IS NOT NULL  AND NM.TX_TYPE = 'I' ) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT BA.KEY_NAME BANK_ACCOUNT_NAME  "
      SelectStr1 = SelectStr1 & ",NM.BANK_ACCOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF BA ON (NM.BANK_ACCOUNT = BA.KEY_ID) "
      
      WhereStr = " WHERE  (NM.BILLING_DOC_ID IS NOT NULL  AND NM.TX_TYPE = 'I' AND NM.BANK_ACCOUNT IS NOT NULL) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT SUM(NM.NET_AMOUNT) NET_AMOUNT "
      SelectStr1 = SelectStr1 & ",NM.BANK_ACCOUNT,NM.BILLING_DOC_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF BA ON (NM.BANK_ACCOUNT = BA.KEY_ID) "
      
      WhereStr = " WHERE  (NM.BILLING_DOC_ID IS NOT NULL  AND NM.TX_TYPE = 'I' AND NM.BANK_ACCOUNT IS NOT NULL) "
   End If
   
   If CASH_TRAN_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_TRAN_ID = " & CASH_TRAN_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
      
   If CASH_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_ID = " & CASH_DOC_ID & ") " & vbCrLf
   End If
      
   If BANK_ACCOUNT > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BANK_ACCOUNT = " & BANK_ACCOUNT & ") " & vbCrLf
   End If
            
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.TX_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.TX_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If Trim(APAR_CODE) <> "" Then
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(APM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(APM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (APM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(BANK_FLAG) <> "" Then
      If WildCard(BANK_FLAG, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CQ.BANK_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CQ.BANK_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CQ.BANK_FLAG = '" & ChangeQuote(Trim(BANK_FLAG)) & "')"
         Else
            WhereStr = WhereStr & "AND (CQ.BANK_FLAG = '" & ChangeQuote(Trim(BANK_FLAG)) & "')"
         End If
      End If
   End If
   
   If Trim(POST_FLAG) <> "" Then
      If WildCard(POST_FLAG, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CQ.POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CQ.POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CQ.POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
         Else
            WhereStr = WhereStr & "AND (CQ.POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
         End If
      End If
   End If
   
   If Trim(TX_TYPE) <> "" Then
      If WildCard(TX_TYPE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.TX_TYPE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.TX_TYPE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "')"
         End If
      End If
   End If
   
   GroupByStr = ""
   If Ind = 4 Then
      GroupByStr = "GROUP BY NM.BANK_ACCOUNT,NM.BILLING_DOC_ID "
   End If
   
   If Ind = 2 Then
      OrderByStr = "ORDER BY CQ.CHEQUE_DATE,CQ.CHEQUE_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY BA.KEY_CODE "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY NM.BILLING_DOC_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CASH_TRAN_ID", NVLI(Rs("CASH_TRAN_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("CASH_DOC_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("CHEQUE_ID", NVLI(Rs("CHEQUE_ID"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("BANK_ACCOUNT", NVLI(Rs("BANK_ACCOUNT"), -1))
      
      Call SetFieldValue("PAYMENT_TYPE", NVLI(Rs("PAYMENT_TYPE"), -1))
      Call SetFieldValue("TX_TYPE", NVLS(Rs("TX_TYPE"), ""))
      Call SetFieldValue("AMOUNT", NVLD(Rs("AMOUNT"), 0))
      Call SetFieldValue("FEE_AMOUNT", NVLD(Rs("FEE_AMOUNT"), 0))
      Call SetFieldValue("NET_AMOUNT", NVLD(Rs("NET_AMOUNT"), 0))
      Call SetFieldValue("TX_NO", NVLS(Rs("TX_NO"), ""))
      Call SetFieldValue("TX_DATE", InternalDateToDate(NVLS(Rs("TX_DATE"), "")))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("STEP_ID", NVLI(Rs("STEP_ID"), -1))
      
      Call SetFieldValue("OLD_CHEQUE_ID", NVLI(Rs("CHEQUE_ID"), -1))
      Call SetFieldValue("OLD_PAYMENT_TYPE", NVLI(Rs("PAYMENT_TYPE"), -1))
      Call SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(NVLI(Rs("PAYMENT_TYPE"), -1)))
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("ACCOUNT_NAME", NVLS(Rs("ACCOUNT_NAME"), ""))
      Call SetFieldValue("CHEQUE_NO", NVLS(Rs("CHEQUE_NO"), ""))
      Call SetFieldValue("CHEQUE_DATE", InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), "")))
      Call SetFieldValue("CHEQUE_AMOUNT", NVLD(Rs("CHEQUE_AMOUNT"), 0))
      
   ElseIf Ind = 2 Then
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      BRANCH_NAME = NVLS(Rs("BRANCH_NAME"), "")
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      CHEQUE_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), ""))
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      CHEQUE_ID = NVLI(Rs("CHEQUE_ID"), -1)
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BANK_BRANCH = NVLI(Rs("BANK_BRANCH"), -1)
   ElseIf Ind = 3 Then
      BANK_ACCOUNT = NVLI(Rs("BANK_ACCOUNT"), -1)
      BANK_ACCOUNT_NAME = NVLS(Rs("BANK_ACCOUNT_NAME"), "")
   ElseIf Ind = 4 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      BANK_ACCOUNT = NVLI(Rs("BANK_ACCOUNT"), -1)
      NET_AMOUNT = NVLD(Rs("NET_AMOUNT"), 0)
   End If
End Sub

Public Sub RegisterFields()
   TableName = "CASH_TRAN"
   SequenceName = "CASH_TRAN_SEQ"
   PKField = "CASH_TRAN_ID"
   
   
      Call AddField("CASH_TRAN_ID", INT_TYPE, -1, ID_CAT)
      Call AddField("BILLING_DOC_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("CHEQUE_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("BANK_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("BANK_BRANCH", INT_TYPE, -1, DATA_CAT)
      Call AddField("TX_TYPE", STRING_TYPE, "", DATA_CAT)
      Call AddField("AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("PAYMENT_TYPE", INT_TYPE, -1, DATA_CAT)
      Call AddField("BANK_ACCOUNT", INT_TYPE, -1, DATA_CAT)
      Call AddField("CASH_DOC_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("FEE_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("NET_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TX_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("TX_DATE", DATE_TYPE, -1, DATA_CAT)
      Call AddField("APAR_MAS_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("EMP_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("STEP_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("BILLING_DOC_PACK", INT_TYPE, -1, DATA_CAT)
      Call AddField("FROM_BILLING", STRING_TYPE, "", DATA_CAT)       'ทำหน้าที่เป็นตัวที่บอกว่าจะเอารายการไหนมาแสดงหรือ SUM เข้าที่ BillingDoc
      
      Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
      Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
      Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
      Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
      Call AddField("OLD_PAYMENT_TYPE", INT_TYPE, -1, TEMP_CAT)
      Call AddField("OLD_CHEQUE_ID", INT_TYPE, -1, TEMP_CAT)
      Call AddField("PAYMENT_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("BANK_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("ACCOUNT_NO", STRING_TYPE, "", TEMP_CAT)
      Call AddField("ACCOUNT_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("CHEQUE_NO", STRING_TYPE, "", TEMP_CAT)
      Call AddField("CHEQUE_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("CHEQUE_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      Call AddField("APAR_CODE", STRING_TYPE, "", TEMP_CAT)
      Call AddField("CUSTOMER_NAME", STRING_TYPE, "", TEMP_CAT)
      
      Call AddField("BANK_FLAG", STRING_TYPE, "", TEMP_CAT)
      Call AddField("POST_FLAG", STRING_TYPE, "", TEMP_CAT)
      
      Call AddField("RECEIPT_TYPE", INT_TYPE, "", TEMP_CAT)
   
      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
      Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   
   
End Sub

Public Sub UpdateBankBranchAccount()
Dim SQL As String
Dim BANK_ID As Long
Dim BANK_BRANCH As Long
Dim BANK_ACCOUNT As Long
Dim CASH_TRAN_ID As Long
Dim TempString1 As String
Dim TempString2 As String
Dim TempString3 As String
   
   BANK_ID = GetFieldValue("BANK_ID")
   BANK_BRANCH = GetFieldValue("BANK_BRANCH")
   BANK_ACCOUNT = GetFieldValue("BANK_ACCOUNT")
   CASH_TRAN_ID = GetFieldValue("CASH_TRAN_ID")
   
   TempString1 = BANK_ID
   If BANK_ID <= 0 Then
      TempString1 = "NULL"
   End If
   TempString2 = BANK_BRANCH
   If BANK_BRANCH <= 0 Then
      TempString2 = "NULL"
   End If
   TempString3 = BANK_ACCOUNT
   If BANK_ACCOUNT <= 0 Then
      TempString3 = "NULL"
   End If
      
   SQL = "UPDATE CASH_TRAN SET BANK_ID = " & TempString1 & ", BANK_BRANCH = " & TempString2 & ", BANK_ACCOUNT = " & TempString3 & " WHERE CASH_TRAN_ID = " & CASH_TRAN_ID & ""
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub DeleteDataBillingDoc()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   If GetFieldValue("CHEQUE_ID") > 0 Then
      SQL = "DELETE FROM " & TableName & " WHERE BILLING_DOC_ID  = " & GetFieldValue("BILLING_DOC_ID") & " AND CHEQUE_ID = " & GetFieldValue("CHEQUE_ID")
   Else
      SQL = "DELETE FROM " & TableName & " WHERE BILLING_DOC_ID  = " & GetFieldValue("BILLING_DOC_ID") & " AND CHEQUE_ID IS NULL "
   End If
   
   Call m_ADOConn.Execute(SQL)
   
End Sub
