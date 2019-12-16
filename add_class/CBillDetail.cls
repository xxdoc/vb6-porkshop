VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

' BILL_DETAIL---------------------------------------------------------------------------------------------------------------------------------
Public BILL_DETAIL_ID As Long
Public SUM_BILL_ID As Long
Public BILLING_DOC_ID As Long
' BILL_DETAIL---------------------------------------------------------------------------------------------------------------------------------
' BILLING_DOC---------------------------------------------------------------------------------------------------------------------------------
Public SUMMARY_DOC_NO As String
Public SUMMARY_DOC_DATE As Date
Public APAR_NAME As String
Public APAR_CODE As String

Public PAID_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public DEBIT_AMOUNT As Double

Public CUSTOMER_BRANCH As Long
Public COUNT_TYPE As Long
Public DOC_ID_TYPE As Long
' BILLING_DOC---------------------------------------------------------------------------------------------------------------------------------
' MASTER_REF
Public CUSTOMER_BRANCH_NAME As String
' MASTER_REF
' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public RcpCnDnItems As Collection

Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_SUMMARY_DOC_DATE As Date
Public TO_SUMMARY_DOC_DATE As Date

' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
Dim iCount As Long
Dim TempRs As ADODB.Recordset
   
   SQL = "DELETE FROM BILL_DETAIL WHERE BILL_DETAIL_ID  = " & BILL_DETAIL_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False, Optional Database As Long = 1)
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
      SelectStr1 = "SELECT BLD.*,BD.DOCUMENT_NO SUMMARY_DOC_NO,BD.DOCUMENT_DATE SUMMARY_DOC_DATE "
      SelectStr1 = SelectStr1 & ",NM2.LONG_NAME APAR_NAME, APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",BD.PAID_AMOUNT,BD.CREDIT_AMOUNT,BD.DEBIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILL_DETAIL BLD "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BLD.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (BD.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
   ElseIf Ind = 2 Then     'CReportNormalSumBill001
      SelectStr1 = "SELECT DISTINCT BD.CUSTOMER_BRANCH,MR2.KEY_NAME CUSTOMER_BRANCH_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILL_DETAIL BLD "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BLD.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (BD.CUSTOMER_BRANCH = MR2.KEY_ID) "
   ElseIf Ind = 3 Then     'GetSumBillByCustomerBranch
      SelectStr1 = "SELECT BD.CUSTOMER_BRANCH,SUM(BD.PAID_AMOUNT) PAID_AMOUNT,SUM(BD.CREDIT_AMOUNT) CREDIT_AMOUNT,SUM(BD.DEBIT_AMOUNT) DEBIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILL_DETAIL BLD "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BLD.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
   ElseIf Ind = 4 Then     'GetCountSumBillByCustomerBranch
      SelectStr1 = "SELECT COUNT (*)  AS COUNT_TYPE,BD2.DOCUMENT_TYPE DOC_ID_TYPE,BD.CUSTOMER_BRANCH "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILL_DETAIL BLD "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BLD.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RCPCNDN_ITEM RCP ON (BD.BILLING_DOC_ID = RCP.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD2 ON (RCP.DOC_ID = BD2.BILLING_DOC_ID) "
   ElseIf Ind = 5 Then  'GetBillDetailReceipted
      SelectStr1 = "SELECT DISTINCT BLD.BILLING_DOC_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM BILL_DETAIL BLD "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BLD.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
   End If
   
   If BILL_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BLD.BILL_DETAIL_ID = " & BILL_DETAIL_ID & ") " & vbCrLf
   End If
   
   If SUM_BILL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BLD.SUM_BILL_ID = " & SUM_BILL_ID & ") " & vbCrLf
   End If
   
'   If FROM_DATE > 0 Then
'      TempDate = DateToStringIntLow(FROM_DATE)
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      WhereStr = WhereStr & TempStr & " (SB.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
'   End If
'
'   If TO_DATE > 0 Then
'      TempDate = DateToStringIntHi(TO_DATE)
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      WhereStr = WhereStr & TempStr & " (SB.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
'   End If
   
   If FROM_SUMMARY_DOC_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_SUMMARY_DOC_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

   If TO_SUMMARY_DOC_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_SUMMARY_DOC_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

   If Ind = 3 Then
      GroupByStr = " GROUP BY BD.CUSTOMER_BRANCH "
   ElseIf Ind = 4 Then
      GroupByStr = " GROUP BY BD2.DOCUMENT_TYPE,BD.CUSTOMER_BRANCH "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY BD.DOCUMENT_DATE ASC,BD.DOCUMENT_NO "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY MR2.KEY_NAME ASC "
   ElseIf Ind = 5 Then
   End If
   
   If Len(OrderByStr) > 0 Then
      If ORDER_TYPE = 1 Then
         OrderByStr = OrderByStr & " ASC "
      ElseIf ORDER_TYPE = 2 Then
         OrderByStr = OrderByStr & " DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
'   Debug.Print SQL1
   If Database = 1 Then
      Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
   ElseIf Database = 2 Then
      Call glbDatabaseMngr.GetRs2(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
   Else
      Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
   End If
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BILL_DETAIL_ID = NVLI(Rs("BILL_DETAIL_ID"), -1)
      SUM_BILL_ID = NVLI(Rs("SUM_BILL_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      
      SUMMARY_DOC_NO = NVLS(Rs("SUMMARY_DOC_NO"), "")
      SUMMARY_DOC_DATE = InternalDateToDate(NVLS(Rs("SUMMARY_DOC_DATE"), ""))
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
   ElseIf Ind = 2 Then
      CUSTOMER_BRANCH = NVLI(Rs("CUSTOMER_BRANCH"), -1)
      CUSTOMER_BRANCH_NAME = NVLS(Rs("CUSTOMER_BRANCH_NAME"), "")
   ElseIf Ind = 3 Then
      CUSTOMER_BRANCH = NVLI(Rs("CUSTOMER_BRANCH"), -1)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      CUSTOMER_BRANCH = NVLI(Rs("CUSTOMER_BRANCH"), -1)
      COUNT_TYPE = NVLI(Rs("COUNT_TYPE"), -1)
      DOC_ID_TYPE = NVLI(Rs("DOC_ID_TYPE"), -1)
   ElseIf Ind = 5 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
   End If
   
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = SUM_BILL_ID
   If SUM_BILL_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = BILLING_DOC_ID
   If BILLING_DOC_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("BILL_DETAIL_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BILL_DETAIL " & _
                     "( " & _
                     "   BILL_DETAIL_ID, " & _
                     "   SUM_BILL_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING BILL_DETAIL_ID "
         BILL_DETAIL_ID = ID
'   Else 'Show edit
'      SQL1 = "UPDATE BILL_DETAIL SET " & _
'                     "    BILL_DETAIL_DESC = '" & ChangeQuote(Trim(BILL_DETAIL_DESC)) & "', " & _
'                     "   DOCUMENT_DATE= '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
'                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
'                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
'                     "WHERE BILL_DETAIL_ID = " & BILL_DETAIL_ID
   End If
   
   If ShowMode = SHOW_ADD Then
      BILL_DETAIL_ID = m_ADOConn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf ShowMode = SHOW_EDIT Then
      Call m_ADOConn.Execute(SQL1)
   End If
End Function
