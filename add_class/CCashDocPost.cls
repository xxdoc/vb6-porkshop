VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCashDocPost"
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

Public CashTranItems As Collection
Public TransferItems As Collection
Public JournalItems As Collection
Public PostItems As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set CashTranItems = New Collection
   Set TransferItems = New Collection
   Set JournalItems = New Collection
   Set PostItems = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set CashTranItems = Nothing
   Set TransferItems = Nothing
   Set JournalItems = Nothing
   Set PostItems = Nothing
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

   Set Tf = m_FieldList(PKField)

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
Dim TempRs As ADODB.Recordset
Dim Ct As CCashTran
Dim iCount As Long
Dim Cq As CCheque

   Set Tf = m_FieldList(PKField)
     
    SQL = "UPDATE CHEQUE  SET  POST_FLAG = 'N'  WHERE  CHEQUE_ID  = " & GetFieldValue("CHEQUE_ID")
   Call m_ADOConn.Execute(SQL)
     
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

Dim CASH_DOC_ID As Long
Dim CASH_DOC_POST_ID As Long
Dim ORDER_TYPE As Long


   CASH_DOC_POST_ID = Val(GetFieldValue("CASH_DOC_POST_ID"))
   CASH_DOC_ID = Val(GetFieldValue("CASH_DOC_ID"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
      
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.KEY_NAME BANK_NAME, MR2.KEY_NAME BRANCH_NAME,CQ.CHEQUE_NO,CQ.CHEQUE_DATE  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.BANK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.BANK_BRANCH = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE CQ ON (NM.CHEQUE_ID = CQ.CHEQUE_ID) "
      
      WhereStr = ""
   End If
   
    If CASH_DOC_POST_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_POST_ID = " & CASH_DOC_POST_ID & ") " & vbCrLf
   End If
   
   If CASH_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_ID = " & CASH_DOC_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY  CASH_DOC_POST_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CASH_DOC_POST_ID", NVLI(Rs("CASH_DOC_POST_ID"), -1))
      Call SetFieldValue("CASH_DOC_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("CHEQUE_ID", NVLI(Rs("CHEQUE_ID"), -1))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("CHEQUE_AMOUNT", NVLD(Rs("CHEQUE_AMOUNT"), 0))
      Call SetFieldValue("CHEQUE_DATE", InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), "")))
      
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("CHEQUE_NO", NVLS(Rs("CHEQUE_NO"), ""))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CASH_DOC_POST"
   SequenceName = "CASH_DOC_POST_SEQ"
   PKField = "CASH_DOC_POST_ID"
   
   Call AddField("CASH_DOC_POST_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("CASH_DOC_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_BRANCH", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CHEQUE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CHEQUE_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("POST_TYPE", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
   'Call AddField("BANK_NO", STRING_TYPE, "", TEMP_CAT)
   'Call AddField("BRANCH_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BANK_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CHEQUE_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CHEQUE_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

   
