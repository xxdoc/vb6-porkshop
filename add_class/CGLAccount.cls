VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGLAccount"
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

Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   
   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   Call m_ADOConn.Execute(SQL)
   
   SQL = "UPDATE GL_ACCOUNT SET CHILD_COUNT = (SELECT COUNT(*) FROM GL_ACCOUNT WHERE PARENT_ID = " & GetFieldValue(PKField) & ") WHERE GL_ACCOUNT_ID = " & GetFieldValue(PKField) & " "
   m_ADOConn.Execute (SQL)

   SQL = "UPDATE GL_ACCOUNT SET CHILD_COUNT = (SELECT COUNT(*) FROM GL_ACCOUNT WHERE PARENT_ID = " & GetFieldValue("PARENT_ID") & ")  WHERE GL_ACCOUNT_ID = " & GetFieldValue("PARENT_ID") & " "
   m_ADOConn.Execute (SQL)
End Sub

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "UPDATE GL_ACCOUNT SET CHILD_COUNT = (SELECT COUNT(*) FROM GL_ACCOUNT WHERE PARENT_ID = " & GetFieldValue(PKField) & ") WHERE GL_ACCOUNT_ID = " & GetFieldValue(PKField) & " "
   m_ADOConn.Execute (SQL)

   SQL = "UPDATE GL_ACCOUNT SET CHILD_COUNT = (SELECT COUNT(*) FROM GL_ACCOUNT WHERE PARENT_ID = " & GetFieldValue("PARENT_ID") & ")  WHERE GL_ACCOUNT_ID = " & GetFieldValue("PARENT_ID") & " "
   m_ADOConn.Execute (SQL)
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

Dim ACC_CODE As String
Dim GL_ACCOUNT_ID As Long

Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   GL_ACCOUNT_ID = Val(GetFieldValue("GL_ACCOUNT_ID"))
   ACC_CODE = GetFieldValue("ACC_CODE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
   
   If GL_ACCOUNT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.GL_ACCOUNT_ID = " & GL_ACCOUNT_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(ACC_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(ACC_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(ACC_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (ACC_CODE = '" & ChangeQuote(Trim(ACC_CODE)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY ACC_CODE "
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
      Call SetFieldValue("GL_ACCOUNT_ID", NVLI(Rs("GL_ACCOUNT_ID"), -1))
      Call SetFieldValue("ACC_CODE", NVLS(Rs("ACC_CODE"), ""))
      Call SetFieldValue("ACC_NAME", NVLS(Rs("ACC_NAME"), ""))
      Call SetFieldValue("ACC_DESC", NVLS(Rs("ACC_DESC"), ""))
      Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), -1))
      Call SetFieldValue("CHILD_COUNT", NVLI(Rs("CHILD_COUNT"), 0))
      
   End If
End Sub

Public Sub RegisterFields()
   TableName = "GL_ACCOUNT"
   SequenceName = "GL_ACCOUNT_SEQ"
   PKField = "GL_ACCOUNT_ID"
   
   Call AddField("GL_ACCOUNT_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("ACC_CODE", STRING_TYPE, "", DATA_CAT)
   Call AddField("ACC_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("ACC_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("PARENT_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("CHILD_COUNT", INT_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

