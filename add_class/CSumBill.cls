VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSumBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

' SUM_BILL---------------------------------------------------------------------------------------------------------------------------------
Public SUM_BILL_ID As Long
Public DOCUMENT_DATE As Date
Public SUM_BILL_DESC As String
' SUM_BILL---------------------------------------------------------------------------------------------------------------------------------
' BILL_DETAIL---------------------------------------------------------------------------------------------------------------------------------
' BILL_DETAIL---------------------------------------------------------------------------------------------------------------------------------
' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public RECEIPT_DOC_NO As String


Public PAID_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public DEBIT_AMOUNT As Double

Public FROM_DATE As Date
Public TO_DATE As Date
' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Public RcpCnDnItems As Collection

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set RcpCnDnItems = New Collection
End Sub
Private Sub Class_Terminate()
   Set RcpCnDnItems = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
Dim iCount As Long
Dim TempRs As ADODB.Recordset
   
   SQL = "DELETE FROM BILL_DETAIL WHERE SUM_BILL_ID  = " & SUM_BILL_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM SUM_BILL WHERE SUM_BILL_ID  = " & SUM_BILL_ID
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
      SelectStr1 = "SELECT SB.* "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUM_BILL SB "
      
   End If
   
   If SUM_BILL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SB.SUM_BILL_ID = " & SUM_BILL_ID & ") " & vbCrLf
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SB.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SB.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(Trim(RECEIPT_DOC_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(RECEIPT_DOC_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM BILL_DETAIL RCP LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = RCP.BILLING_DOC_ID) WHERE (RCP.SUM_BILL_ID = SB.SUM_BILL_ID) AND (SUBSTR(BD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)"
      Else
         WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM BILL_DETAIL RCP LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = RCP.BILLING_DOC_ID) WHERE (RCP.SUM_BILL_ID = SB.SUM_BILL_ID) AND (BD.DOCUMENT_NO = '" & ChangeQuote(Trim(RECEIPT_DOC_NO)) & "') > 0)"
      End If
   End If
   
   If Ind = 3 Then
      'GroupByStr = " GROUP BY NM.DOCUMENT_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME,NM.DOCUMENT_TYPE "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY SB.DOCUMENT_DATE "
   ElseIf Ind = 2 Then
   ElseIf Ind = 5 Then
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
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
      SUM_BILL_ID = NVLI(Rs("SUM_BILL_ID"), -1)
      SUM_BILL_DESC = NVLS(Rs("SUM_BILL_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
   End If
   
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

'   TempID1 = BILLING_ADDRESS_ID
'   If BILLING_ADDRESS_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
   
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("SUM_BILL_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SUM_BILL " & _
                     "( " & _
                     "   SUM_BILL_ID, " & _
                     "   SUM_BILL_DESC, " & _
                     "   DOCUMENT_DATE, " & _
                     "   PAID_AMOUNT, " & _
                     "   CREDIT_AMOUNT, " & _
                     "   DEBIT_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(SUM_BILL_DESC)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                        PAID_AMOUNT & "," & _
                        CREDIT_AMOUNT & "," & _
                        DEBIT_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") RETURNING SUM_BILL_ID "
         SUM_BILL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUM_BILL SET " & _
                     "    SUM_BILL_DESC = '" & ChangeQuote(Trim(SUM_BILL_DESC)) & "', " & _
                     "   DOCUMENT_DATE= '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "    PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "    CREDIT_AMOUNT = " & CREDIT_AMOUNT & ", " & _
                     "    DEBIT_AMOUNT = " & DEBIT_AMOUNT & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUM_BILL_ID = " & SUM_BILL_ID
   End If
   
   If ShowMode = SHOW_ADD Then
      SUM_BILL_ID = m_ADOConn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf ShowMode = SHOW_EDIT Then
      Call m_ADOConn.Execute(SQL1)
   End If
End Function
