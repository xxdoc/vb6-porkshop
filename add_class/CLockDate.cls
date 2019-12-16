VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLockDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public LOCK_DATE_ID As Long
Public FROM_DATE As Date
Public TO_DATE As Date

Public FROM_INVENTORY_DATE As Date
Public TO_INVENTORY_DATE As Date

Public FROM_INVOICE_DATE As Date
Public TO_INVOICE_DATE As Date

Public FROM_RECEIPT_DATE As Date
Public TO_RECEIPT_DATE As Date

Public LOCK_TYPE As Long
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
   
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
   
   If Ind = 1 Then
      SelectStr1 = "SELECT LD.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM LOCK_DATE LD "
      
      WhereStr = ""
   End If
   
   If LOCK_DATE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (LD.LOCK_DATE_ID = " & LOCK_DATE_ID & ") " & vbCrLf
   End If
   
   If LOCK_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (LD.LOCK_TYPE = " & LOCK_TYPE & ") " & vbCrLf
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      LOCK_DATE_ID = NVLI(Rs("LOCK_DATE_ID"), -1)
      FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
      TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
      LOCK_TYPE = NVLI(Rs("LOCK_TYPE"), -1)
      
      FROM_INVENTORY_DATE = InternalDateToDate(NVLS(Rs("FROM_INVENTORY_DATE"), ""))
      TO_INVENTORY_DATE = InternalDateToDate(NVLS(Rs("TO_INVENTORY_DATE"), ""))
      
      FROM_INVOICE_DATE = InternalDateToDate(NVLS(Rs("FROM_INVOICE_DATE"), ""))
      TO_INVOICE_DATE = InternalDateToDate(NVLS(Rs("TO_INVOICE_DATE"), ""))
      
      FROM_RECEIPT_DATE = InternalDateToDate(NVLS(Rs("FROM_RECEIPT_DATE"), ""))
      TO_RECEIPT_DATE = InternalDateToDate(NVLS(Rs("TO_RECEIPT_DATE"), ""))
      
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE LOCK_DATE SET " & _
                     "   FROM_DATE= '" & ChangeQuote(Trim(DateToStringInt(FROM_DATE))) & "', " & _
                     "   TO_DATE= '" & ChangeQuote(Trim(DateToStringInt(TO_DATE))) & "', " & _
                     "   FROM_INVENTORY_DATE= '" & ChangeQuote(Trim(DateToStringInt(FROM_INVENTORY_DATE))) & "', " & _
                     "   TO_INVENTORY_DATE= '" & ChangeQuote(Trim(DateToStringInt(TO_INVENTORY_DATE))) & "', " & _
                     "   FROM_INVOICE_DATE= '" & ChangeQuote(Trim(DateToStringInt(FROM_INVOICE_DATE))) & "', " & _
                     "   TO_INVOICE_DATE= '" & ChangeQuote(Trim(DateToStringInt(TO_INVOICE_DATE))) & "', " & _
                     "   FROM_RECEIPT_DATE= '" & ChangeQuote(Trim(DateToStringInt(FROM_RECEIPT_DATE))) & "', " & _
                     "   TO_RECEIPT_DATE= '" & ChangeQuote(Trim(DateToStringInt(TO_RECEIPT_DATE))) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOCK_DATE_ID = " & LOCK_DATE_ID
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
