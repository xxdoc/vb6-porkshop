VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJournal"
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

Public JournalItems As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   Set JournalItems = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set JournalItems = Nothing
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
   
   SQL = "DELETE FROM JOURNAL_ITEM WHERE " & PKField & " = " & Tf.GetValue
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

Dim JOURNAL_NO As String
Dim JOURNAL_ID As Long
Dim JOURNAL_TYPE As Long

Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim POST_FLAG As String
Dim APAR_NAME As String
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   JOURNAL_ID = Val(GetFieldValue("JOURNAL_ID"))
   JOURNAL_TYPE = Val(GetFieldValue("JOURNAL_TYPE"))
   JOURNAL_NO = GetFieldValue("JOURNAL_NO")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   APAR_NAME = GetFieldValue("APAR_NAME")
   POST_FLAG = GetFieldValue("POST_FLAG")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, NM2.LONG_NAME APAR_NAME, APM.APAR_IND "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      WhereStr = ""
   End If
   
   If JOURNAL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.JOURNAL_ID = " & JOURNAL_ID & ") " & vbCrLf
   End If
   
   If JOURNAL_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (NM.JOURNAL_TYPE = " & JOURNAL_TYPE & ") " & vbCrLf
   End If
   
   If Len(Trim(JOURNAL_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(JOURNAL_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(JOURNAL_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (JOURNAL_NO = '" & ChangeQuote(Trim(JOURNAL_NO)) & "')"
      End If
   End If
   
   If Len(Trim(POST_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(POST_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(APAR_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM2.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM2.LONG_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (JOURNAL_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (JOURNAL_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY JOURNAL_NO "
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
      Call SetFieldValue("JOURNAL_ID", NVLI(Rs("JOURNAL_ID"), -1))
      Call SetFieldValue("JOURNAL_NO", NVLS(Rs("JOURNAL_NO"), ""))
      Call SetFieldValue("JOURNAL_DATE", InternalDateToDate(NVLS(Rs("JOURNAL_DATE"), "")))
      Call SetFieldValue("JOURNAL_TYPE", NVLI(Rs("JOURNAL_TYPE"), -1))
      Call SetFieldValue("JOURNAL_DESC", NVLS(Rs("JOURNAL_DESC"), ""))
      Call SetFieldValue("JOURNAL_AMOUNT", NVLI(Rs("JOURNAL_AMOUNT"), 0))
      Call SetFieldValue("POST_FLAG", NVLS(Rs("POST_FLAG"), ""))
      Call SetFieldValue("POST_DATE", InternalDateToDate(NVLS(Rs("POST_DATE"), "")))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("DEPARTMENT_ID", NVLI(Rs("DEPARTMENT_ID"), -1))
   
      Call SetFieldValue("APAR_NAME", NVLS(Rs("APAR_NAME"), ""))
      Call SetFieldValue("APAR_IND", NVLI(Rs("APAR_IND"), -1))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "JOURNAL"
   SequenceName = "JOURNAL_SEQ"
   PKField = "JOURNAL_ID"
   
   Call AddField("JOURNAL_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("JOURNAL_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("JOURNAL_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("JOURNAL_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("JOURNAL_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("JOURNAL_AMOUNT", MONEY_TYPE, "", DATA_CAT)
   Call AddField("POST_FLAG", STRING_TYPE, "", DATA_CAT)
   Call AddField("POST_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("APAR_MAS_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("DEPARTMENT_ID", INT_TYPE, "", DATA_CAT)

   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("APAR_IND", INT_TYPE, "", TEMP_CAT)
   Call AddField("APAR_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub

