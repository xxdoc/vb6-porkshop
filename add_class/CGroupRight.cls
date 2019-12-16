VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupRight"
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

Dim GROUP_RIGHT_ID As Long
Dim GROUP_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   GROUP_RIGHT_ID = Val(GetFieldValue("GROUP_RIGHT_ID"))
   GROUP_ID = Val(GetFieldValue("GROUP_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT RI.*, EN.*, EN.RIGHT_ID RIGHT_ID2, EN.RIGHT_STATUS RIGHT_STATUS2, RI.RIGHT_ID RIGHT_ID1, RI.DEFAULT_VALUE RIGHT_STATUS1 "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM RIGHT_ITEM RI "
      FromStr = FromStr & "LEFT OUTER JOIN GROUP_RIGHT EN ON ((RI.RIGHT_ID = EN.RIGHT_ID) AND (EN.GROUP_ID = " & GROUP_ID & "))"
   
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT EN.*, EN.RIGHT_ID RIGHT_ID_ORG, RI.RIGHT_ITEM_NAME, RI.DEFAULT_VALUE  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM GROUP_RIGHT EN "
      FromStr = FromStr & "LEFT OUTER JOIN RIGHT_ITEM RI ON (EN.RIGHT_ID = RI.RIGHT_ID) "
   End If
   
   If GROUP_RIGHT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EN.GROUP_RIGHT_ID = " & GROUP_RIGHT_ID & ") " & vbCrLf
   End If
   
   If (GROUP_ID > 0) And (Ind <> 2) Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EN.GROUP_ID = " & GROUP_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY GROUP_ID "
   ElseIf Ind = 2 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY RI.RIGHT_ITEM_NAME "
      Else
         OrderByStr = "ORDER BY RI.RIGHT_ITEM_NAME "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY GROUP_ID "
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
      Call SetFieldValue("GROUP_RIGHT_ID", NVLI(Rs("GROUP_RIGHT_ID"), -1))
      Call SetFieldValue("GROUP_ID", NVLI(Rs("GROUP_ID"), -1))
      Call SetFieldValue("RIGHT_ID", NVLI(Rs("RIGHT_ID"), -1))
      Call SetFieldValue("RIGHT_STATUS", NVLS(Rs("RIGHT_STATUS"), ""))
   ElseIf Ind = 2 Then
         Call SetFieldValue("GROUP_ID", NVLI(Rs("GROUP_ID"), -1))
         Call SetFieldValue("GROUP_RIGHT_ID", NVLI(Rs("GROUP_RIGHT_ID"), -1))
         Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), 0))
         If NVLI(Rs("RIGHT_ID1"), 0) <> NVLI(Rs("RIGHT_ID2"), 0) Then
            Call SetFieldValue("RIGHT_ID", NVLI(Rs("RIGHT_ID1"), 0))
            Call SetFieldValue("RIGHT_STATUS", NVLS(Rs("RIGHT_STATUS1"), ""))
            Flag = "A"
         Else
            Call SetFieldValue("RIGHT_ID", NVLI(Rs("RIGHT_ID2"), 0))
            Call SetFieldValue("RIGHT_STATUS", NVLS(Rs("RIGHT_STATUS2"), ""))
            Flag = "I"
         End If
         Call SetFieldValue("RIGHT_ITEM_DESC", NVLS(Rs("RIGHT_ITEM_DESC"), ""))
         Call SetFieldValue("RIGHT_ITEM_NAME", NVLS(Rs("RIGHT_ITEM_NAME"), ""))
   ElseIf Ind = 3 Then
      Call SetFieldValue("GROUP_RIGHT_ID", NVLI(Rs("GROUP_RIGHT_ID"), -1))
      Call SetFieldValue("GROUP_ID", NVLI(Rs("GROUP_ID"), -1))
      Call SetFieldValue("RIGHT_ID", NVLI(Rs("RIGHT_ID_ORG"), -1))
'      If GetFieldValue("RIGHT_ID") = 673 Then
'         Debug.Print
'      End If
      Call SetFieldValue("RIGHT_STATUS", NVLS(Rs("RIGHT_STATUS"), ""))
      If GetFieldValue("RIGHT_ID") <= 0 Then
         Call SetFieldValue("RIGHT_STATUS", "Y")
      End If
      Call SetFieldValue("RIGHT_ITEM_NAME", NVLS(Rs("RIGHT_ITEM_NAME"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "GROUP_RIGHT"
   SequenceName = "GROUP_RIGHT_SEQ"
   PKField = "GROUP_RIGHT_ID"
   
   Call AddField("GROUP_RIGHT_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("GROUP_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("RIGHT_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("RIGHT_STATUS", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("PARENT_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("RIGHT_ITEM_DESC", STRING_TYPE, "", TEMP_CAT)
   Call AddField("RIGHT_ITEM_NAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

