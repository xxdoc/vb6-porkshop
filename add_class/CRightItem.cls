VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRightItem"
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

Private Sub AddField(FRIGHT_ITEM As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FRIGHT_ITEM, FType, FValue, FCat)
   Call m_FieldList.add(F, FRIGHT_ITEM)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldRIGHT_ITEM As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldRIGHT_ITEM)
   GetFieldValue = Tf.GetValue
End Function

Public Sub SetFieldValue(FieldRIGHT_ITEM As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldRIGHT_ITEM)
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
Dim NewStr As String
Dim SubLen As Long

Dim RIGHT_ITEM_NAME As String
Dim RIGHT_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   RIGHT_ID = Val(GetFieldValue("RIGHT_ID"))
   RIGHT_ITEM_NAME = GetFieldValue("RIGHT_ITEM_NAME")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
   
   If RIGHT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.RIGHT_ID = " & RIGHT_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(RIGHT_ITEM_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(RIGHT_ITEM_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(RIGHT_ITEM_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (RIGHT_ITEM_NAME = '" & ChangeQuote(Trim(RIGHT_ITEM_NAME)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY RIGHT_ITEM_DESC "
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
      Call SetFieldValue("RIGHT_ID", NVLI(Rs("RIGHT_ID"), -1))
      Call SetFieldValue("RIGHT_ITEM_DESC", NVLS(Rs("RIGHT_ITEM_DESC"), ""))
      Call SetFieldValue("DEFAULT_VALUE", NVLS(Rs("DEFAULT_VALUE"), ""))
      Call SetFieldValue("RIGHT_ITEM_NAME", NVLS(Rs("RIGHT_ITEM_NAME"), ""))
      Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), -1))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "RIGHT_ITEM"
   SequenceName = "RIGHT_ITEM_SEQ"
   PKField = "RIGHT_ID"
   
   Call AddField("RIGHT_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("RIGHT_ITEM_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("DEFAULT_VALUE", STRING_TYPE, "", DATA_CAT)
   Call AddField("RIGHT_ITEM_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("PARENT_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("ADDEDIT", INT_TYPE, "1", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub


