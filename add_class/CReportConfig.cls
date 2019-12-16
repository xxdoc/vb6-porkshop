VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String
Public ShowMode As SHOW_MODE_TYPE

Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection
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
Public Sub AddEditData(ShowMode As SHOW_MODE_TYPE)
Dim SQL As String

   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   ''debug.print (SQL)
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
Dim SubLen As Long
Dim NewStr As String

Dim REPORT_CONFIG_ID As Long
Dim PAPER_SIZE As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim REPORT_KEY As String

   REPORT_CONFIG_ID = GetFieldValue("REPORT_CONFIG_ID")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   REPORT_KEY = GetFieldValue("REPORT_KEY")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT MR.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " MR "
      
      WhereStr = ""
   End If
   
   If REPORT_CONFIG_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (MR.REPORT_CONFIG_ID = " & REPORT_CONFIG_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(REPORT_KEY)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(REPORT_KEY, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(REPORT_KEY,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (REPORT_KEY = '" & ChangeQuote(Trim(REPORT_KEY)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY REPORT_KEY "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   ''debug.print (SQL1)
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("REPORT_CONFIG_ID", NVLI(Rs("REPORT_CONFIG_ID"), -1))
      Call SetFieldValue("REPORT_KEY", NVLS(Rs("REPORT_KEY"), ""))
      Call SetFieldValue("PAPER_SIZE", NVLD(Rs("PAPER_SIZE"), 0))
      Call SetFieldValue("ORIENTATION", NVLD(Rs("ORIENTATION"), 0))
      Call SetFieldValue("MARGIN_BOTTOM", NVLD(Rs("MARGIN_BOTTOM"), 0))
      Call SetFieldValue("MARGIN_LEFT", NVLD(Rs("MARGIN_LEFT"), 0))
      Call SetFieldValue("MARGIN_RIGHT", NVLD(Rs("MARGIN_RIGHT"), 0))
      Call SetFieldValue("MARGIN_TOP", NVLD(Rs("MARGIN_TOP"), 0))
      Call SetFieldValue("FONT_NAME", NVLS(Rs("FONT_NAME"), ""))
      Call SetFieldValue("FONT_SIZE", NVLD(Rs("FONT_SIZE"), 0))
      Call SetFieldValue("HEAD_OFFSET", NVLD(Rs("HEAD_OFFSET"), 0))
      Call SetFieldValue("DUMMY_OFFSET", NVLD(Rs("DUMMY_OFFSET"), 0))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "REPORT_CONFIG"
   SequenceName = "REPORT_CONFIG_SEQ"
   PKField = "REPORT_CONFIG_ID"
   
   Call AddField("REPORT_CONFIG_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("REPORT_KEY", STRING_TYPE, "", DATA_CAT)
   Call AddField("PAPER_SIZE", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("ORIENTATION", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("MARGIN_BOTTOM", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("MARGIN_LEFT", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("MARGIN_RIGHT", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("MARGIN_TOP", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("FONT_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("FONT_SIZE", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("HEAD_OFFSET", MONEY_TYPE, "0", DATA_CAT)
   Call AddField("DUMMY_OFFSET", MONEY_TYPE, "0", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

