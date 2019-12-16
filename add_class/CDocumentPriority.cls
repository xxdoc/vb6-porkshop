VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentPriority"
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
Dim TempDate As String

Dim DOCUMENT_PRIORITY_ID As Long
Dim DOCUMENT_TYPE As Long
Dim Area As Long
Dim PRIORITY1 As Long
Dim PRIORITY2 As Long

Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
'

   DOCUMENT_PRIORITY_ID = Val(GetFieldValue("DOCUMENT_PRIORITY_ID"))
   DOCUMENT_TYPE = Val(GetFieldValue("DOCUMENT_TYPE"))
   Area = Val(GetFieldValue("AREA"))
   PRIORITY1 = Val(GetFieldValue("PRIORITY1"))
   PRIORITY2 = Val(GetFieldValue("PRIORITY2"))
   
   
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   '
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
  
   End If
   
   If DOCUMENT_PRIORITY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_PRIORITY_ID = " & DOCUMENT_PRIORITY_ID & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If Area > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.AREA = " & Area & ") " & vbCrLf
   End If

   If PRIORITY1 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PRIORITY1 = " & PRIORITY1 & ") " & vbCrLf
   End If
   
   If PRIORITY2 > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PRIORITY2 = " & PRIORITY2 & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_PRIORITY_ID "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY DOCUMENT_PRIORITY_ID "
      Else
         OrderByStr = "ORDER BY DOCUMENT_PRIORITY_ID "
      End If
   
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
      Call SetFieldValue("DOCUMENT_PRIORITY_ID", NVLI(Rs("DOCUMENT_PRIORITY_ID"), -1))
      Call SetFieldValue("DOCUMENT_TYPE", NVLI(Rs("DOCUMENT_TYPE"), -1))
      Call SetFieldValue("AREA", NVLI(Rs("AREA"), -1))
      Call SetFieldValue("PRIORITY1", NVLI(Rs("PRIORITY1"), -1))
      Call SetFieldValue("PRIORITY2", NVLI(Rs("PRIORITY2"), -1))
      Call SetFieldValue("COMMENT", NVLS(Rs("COMMENT"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "DOCUMENT_PRIORITY"
   SequenceName = "DOCUMENT_PRIORITY_SEQ"
   PKField = "DOCUMENT_PRIORITY_ID"
   
   Call AddField("DOCUMENT_PRIORITY_ID", INT_TYPE, -1, ID_CAT)
   
   Call AddField("DOCUMENT_TYPE", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("AREA", INT_TYPE, -1, DATA_CAT)
   Call AddField("PRIORITY1", INT_TYPE, -1, DATA_CAT)
   Call AddField("PRIORITY2", INT_TYPE, -1, DATA_CAT)
   Call AddField("COMMENT", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
   
End Sub



