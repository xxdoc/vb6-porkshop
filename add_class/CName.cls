VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CName"
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

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
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

Dim NAME_ID As Long
Dim PREFIX_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   NAME_ID = Val(GetFieldValue("NAME_ID"))
   PREFIX_ID = Val(GetFieldValue("PREFIX_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
   
   If NAME_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.NAME_ID = " & NAME_ID & ") " & vbCrLf
   End If
   
   If PREFIX_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.PREFIX_ID = " & PREFIX_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY LONG_NAME "
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
      Call SetFieldValue("NAME_ID", NVLI(Rs("NAME_ID"), -1))
      Call SetFieldValue("LONG_NAME", NVLS(Rs("LONG_NAME"), ""))
      Call SetFieldValue("SHORT_NAME", NVLS(Rs("SHORT_NAME"), ""))
      Call SetFieldValue("PREFIX_ID", NVLI(Rs("PREFIX_ID"), -1))
      Call SetFieldValue("NICK_NAME", NVLS(Rs("NICK_NAME"), ""))
      Call SetFieldValue("MIDDLE_NAME", NVLS(Rs("MIDDLE_NAME"), ""))
      Call SetFieldValue("LAST_NAME", NVLS(Rs("LAST_NAME"), ""))
      Call SetFieldValue("EMAIL", NVLS(Rs("EMAIL"), ""))
      Call SetFieldValue("LANGUAGE_ID", NVLI(Rs("LANGUAGE_ID"), -1))
      Call SetFieldValue("BILL_NAME", NVLS(Rs("BILL_NAME"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "NAME"
   SequenceName = "NAME_SEQ"
   PKField = "NAME_ID"
   
   Call AddField("NAME_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PREFIX_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("LONG_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("SHORT_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("NICK_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("MIDDLE_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("LAST_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("EMAIL", STRING_TYPE, "", DATA_CAT)
   Call AddField("LANGUAGE_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("BILL_NAME", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

