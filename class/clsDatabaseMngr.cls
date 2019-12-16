VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabaseMngr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "clsDatabaseMngr"

Private m_ADOConn As ADODB.Connection
Private m_ADO2Conn As ADODB.Connection
Private m_Rs As ADODB.Recordset

Private Type TProvince
   PROVINCE_ID As Long
   PROVINCE_NAME As String
End Type

Private Type TExpenseType
   EXPENSE_TYPE_ID As Long
   EXPENSE_TYPE_NAME As String
End Type

Private m_Province() As TProvince
Private m_ExpenseType() As TExpenseType

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
Dim RName As String

   RName = "Class_Initialize"
   
   Set m_ADOConn = New ADODB.Connection
   Set m_ADO2Conn = New ADODB.Connection
   Set m_Rs = New ADODB.Recordset
   
   Exit Sub

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Property Get DBConnection() As ADODB.Connection
   Set DBConnection = m_ADOConn
End Property
Public Property Get DB2Connection() As ADODB.Connection
   Set DB2Connection = m_ADO2Conn
End Property
Public Function GetServerDateTime(ServerDateTime As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim ItemCount As Long

   RName = "GetServerDateTime"
   GetServerDateTime = False
   
   SQL = "SELECT CURRENT_TIMESTAMP || '' FROM SYSTEM_PARAM WHERE PARAM_NAME = 'DUMMY_PARAM' "
   If Not GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "End of file condition."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   If NVLS(m_Rs(0), "") = "" Then
      ErrorObj.LocalErrorMsg = "TIMESTAMP is null."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   ServerDateTime = m_Rs(0)
   
   GetServerDateTime = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL -> " & SQL
   ErrorObj.SystemErrorMsg = err.Description
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetServerDateTime = False
End Function

Public Function SetSeqID(SequenceName As String, ID As Long)
Dim SQL As String

   SQL = "SET GENERATOR " & SequenceName & " TO " & ID
   m_ADOConn.Execute (SQL)
End Function
Public Function GetSeqID(SequenceName As String, ID As Long, ErrorObj As clsErrorLog, Optional Space As Long = 1) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim ItemCount As Long

   RName = "GetSeqID"
   GetSeqID = False
   
   SQL = "SELECT GEN_ID(" & SequenceName & ", " & Space & ") FROM SYSTEM_PARAM WHERE PARAM_NAME = 'DUMMY_PARAM' "
   If Not GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "End of file condition."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   If NVLI(m_Rs(0), 0) = 0 Then
      ErrorObj.LocalErrorMsg = "ID is null."
      ErrorObj.SystemErrorMsg = SQL
      Exit Function
   End If
   
   ID = m_Rs(0)
   
   GetSeqID = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL -> " & SQL
   ErrorObj.SystemErrorMsg = err.Description
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetSeqID = False
End Function
Public Function GetRs(SQL1 As String, SQL2 As String, Optional MustCount As Boolean = True, Optional ItemCount As Long, Optional Rs As ADODB.Recordset, Optional ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
      
   RName = "GetRs"
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = err.Description
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   If MustCount Then
      Call Rs.Open(SQL2, m_ADOConn, adOpenKeyset, adLockOptimistic)
      If Rs.EOF Then
         ItemCount = 0
      Else
         ItemCount = NVLI(Rs(0), 0)
      End If
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   
   m_ADOConn.CursorLocation = adUseClient
   Call Rs.Open(SQL1, m_ADOConn, adOpenKeyset, adLockOptimistic)
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = err.Description
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   If Rs.State <> adStateOpen Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = err.Description
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs = False
      Exit Function
   End If
   
   GetRs = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
   ErrorObj.SystemErrorMsg = err.Description
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetRs = False
End Function
Public Function GetRs2(SQL1 As String, SQL2 As String, Optional MustCount As Boolean = True, Optional ItemCount As Long, Optional Rs As ADODB.Recordset, Optional ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
      
   RName = "GetRs2"
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = err.Description
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs2 = False
      Exit Function
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   If MustCount Then
      Call Rs.Open(SQL2, m_ADO2Conn, adOpenKeyset, adLockOptimistic)
      If Rs.EOF Then
         ItemCount = 0
      Else
         ItemCount = NVLI(Rs(0), 0)
      End If
   End If
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   
   m_ADO2Conn.CursorLocation = adUseClient
   Call Rs.Open(SQL1, m_ADO2Conn, adOpenKeyset, adLockOptimistic)
   
   If Rs Is Nothing Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = err.Description
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs2 = False
      Exit Function
   End If
   
   If Rs.State <> adStateOpen Then
      ErrorObj.RoutineName = RName
      ErrorObj.ModuleName = MODULE_NAME
      ErrorObj.LocalErrorMsg = "Rs is null before open are not allow. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
      ErrorObj.SystemErrorMsg = err.Description
      Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
      
      GetRs2 = False
      Exit Function
   End If
   
   GetRs2 = True
   Exit Function
   
ErrorHandler:
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.LocalErrorMsg = "Run time error. Can't get record set. SQL1 -> " & SQL1 & " SQL2 -> " & SQL2
   ErrorObj.SystemErrorMsg = err.Description
   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   
   GetRs2 = False
End Function
Public Function ConnectDatabase(FileDb As String, UserName As String, Password As String, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim ConnStr As String
   
   RName = "ConnectDatabase"
   
   ConnStr = "data source=" & FileDb & ";ctype=UNICODE_FSS"
'   m_ADOConn.Open ("DSN=TestFireBird;")
   m_ADOConn.Provider = "LCPI.IBProvider"
   Call m_ADOConn.Open(ConnStr, UserName, Password)
   
   m_ADOConn.Properties("Session AutoCommit") = True   ' enable auto commit
   m_ADOConn.Properties("Autocommit Isolation Levels") = 4096  ' read commited=0x1000
      
   ConnectDatabase = True
   Exit Function

ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = MODULE_NAME
   Call glbErrorLog.ShowErrorLog(LOG_FILE_MSGBOX, False)

'   Unload frmSplash
'   Set frmSplash = Nothing
   
   ConnectDatabase = False
End Function
Public Function CountRecord(TableName As String, FieldName1 As String, FieldName2 As String, FieldValue As String, ID As Long, ErrorObj As clsErrorLog, Optional FieldNameExTend As String, Optional FieldNameExTendValue As String, Optional FieldNameExTendEX As String, Optional FieldNameExTendValueEX As String, Optional NullFlag As Boolean = False) As Long
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim ItemCount As Long

   RName = "CountRecord"
   CountRecord = False
    
    SQL1 = "SELECT COUNT(*) RC FROM " & TableName & " WHERE (" & FieldName1 & " = '" & FieldValue & "') AND (" & FieldName2 & " <> " & ID & ") "
    
    If Len(FieldNameExTendValue) > 0 Then
      SQL1 = SQL1 & " AND (" & FieldNameExTend & " = '" & FieldNameExTendValue & "')"
   ElseIf NullFlag Then
      SQL1 = SQL1 & " AND (" & FieldNameExTend & " IS NULL)"
   End If
   If Len(FieldNameExTendValueEX) > 0 Then
      SQL1 = SQL1 & " AND (" & FieldNameExTendEX & " = '" & FieldNameExTendValueEX & "')"
   ElseIf NullFlag Then
      SQL1 = SQL1 & " AND (" & FieldNameExTendEX & " IS NULL)"
   End If
   
   If Not GetRs(SQL1, "", False, ItemCount, m_Rs, ErrorObj) Then
      Exit Function
   End If
   
   If m_Rs.EOF Then
      CountRecord = 0
   Else
      CountRecord = NVLI(m_Rs("RC"), 0)
   End If

   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = MODULE_NAME
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   CountRecord = False
End Function

