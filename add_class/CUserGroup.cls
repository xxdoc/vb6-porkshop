VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserGroup"
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
Public RightItems As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set RightItems = New Collection
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set RightItems = Nothing
End Sub

Private Sub AddField(FUSER_GROUP As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FUSER_GROUP, FType, FValue, FCat)
   Call m_FieldList.add(F, FUSER_GROUP)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldUSER_GROUP As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldUSER_GROUP)
   GetFieldValue = Tf.GetValue
End Function

Public Sub SetFieldValue(FieldUSER_GROUP As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldUSER_GROUP)
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
   
   SQL = "DELETE FROM GROUP_RIGHT WHERE " & PKField & " = " & Tf.GetValue
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
Dim SubLen As Long
Dim NewStr As String

Dim GROUP_ID As Long
Dim GROUP_NAME As String
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   GROUP_ID = Val(GetFieldValue("GROUP_ID"))
   GROUP_NAME = GetFieldValue("GROUP_NAME")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
   
   If GROUP_ID >= 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.GROUP_ID = " & GROUP_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(GROUP_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(GROUP_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(GROUP_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (GROUP_NAME = '" & ChangeQuote(Trim(GROUP_NAME)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY GROUP_NAME "
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
      Call SetFieldValue("GROUP_ID", NVLI(Rs("GROUP_ID"), -1))
      Call SetFieldValue("GROUP_NAME", NVLS(Rs("GROUP_NAME"), ""))
      Call SetFieldValue("GROUP_DESC", NVLS(Rs("GROUP_DESC"), ""))
      Call SetFieldValue("GROUP_STATUS", NVLS(Rs("GROUP_STATUS"), ""))
      Call SetFieldValue("MAX_USER", NVLI(Rs("MAX_USER"), -1))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "USER_GROUP"
   SequenceName = "USER_GROUP_SEQ"
   PKField = "GROUP_ID"
   
   Call AddField("GROUP_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("GROUP_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("GROUP_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("GROUP_STATUS", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("MAX_USER", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

