VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatchDelivery"
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
'Private m_ADOConn2 As ADODB.Connection
Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   'Set m_ADOConn2 = glbDatabaseMngr.DB2Connection
   
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

Public Sub AddEditData(ShowMode As SHOW_MODE_TYPE)
Dim SQL As String

   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   Call m_ADOConn.Execute(SQL)
   'Call m_ADOConn2.Execute(SQL)
End Sub

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   'Call m_ADOConn2.Execute(SQL)
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

Dim PATCH_DELIVERY_ID As Long
Dim PATCH_BY As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   PATCH_DELIVERY_ID = GetFieldValue("PATCH_DELIVERY_ID")
   PATCH_BY = GetFieldValue("PATCH_BY")
   ORDER_BY = GetFieldValue("ORDER_BY")
   ORDER_TYPE = GetFieldValue("ORDER_TYPE")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT MR.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " MR "
      
      WhereStr = ""
   End If
   
   If PATCH_DELIVERY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.PATCH_DELIVERY_ID = " & PATCH_DELIVERY_ID & ") " & vbCrLf
   End If
   
   If PATCH_BY > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.PATCH_BY = " & PATCH_BY & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY PATCH_VERSION "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("PATCH_DELIVERY_ID", NVLI(Rs("PATCH_DELIVERY_ID"), -1))
      Call SetFieldValue("PATCH_VERSION", NVLS(Rs("PATCH_VERSION"), ""))
      Call SetFieldValue("PATCH_DATE", InternalDateToDate(NVLS(Rs("PATCH_DATE"), "")))
      Call SetFieldValue("PATCH_BY", NVLI(Rs("PATCH_BY"), -1))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "PATCH_DELIVERY"
   SequenceName = "PATCH_DELIVERY_SEQ"
   PKField = "PATCH_DELIVERY_ID"
   
   Call AddField("PATCH_DELIVERY_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PATCH_VERSION", STRING_TYPE, "", DATA_CAT)
   Call AddField("PATCH_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("PATCH_BY", INT_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

