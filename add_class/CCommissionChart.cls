VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionChart"
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
Dim TempDate As String

Dim COMMISSION_CHART_ID As Long
Dim MASTER_FROMTO_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim FROM_COMMISSION_DATE As Date
Dim TO_COMMISSION_DATE As Date

   COMMISSION_CHART_ID = Val(GetFieldValue("COMMISSION_CHART_ID"))
   MASTER_FROMTO_ID = Val(GetFieldValue("MASTER_FROMTO_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   FROM_COMMISSION_DATE = GetFieldValue("FROM_COMMISSION_DATE")
   TO_COMMISSION_DATE = GetFieldValue("TO_COMMISSION_DATE")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME EMP_NAME, NM3.LAST_NAME EMP_LNAME "
      SelectStr1 = SelectStr1 & ",FT1.MASTER_FROMTO_DESC SALE_NAME, FT2.MASTER_FROMTO_DESC RETURN_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EN.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO FT1 ON (FT1.MASTER_FROMTO_ID = EN.SALE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO FT2 ON (FT2.MASTER_FROMTO_ID = EN.RETURN_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME EMP_NAME, NM3.LAST_NAME EMP_LNAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EN.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME EMP_NAME, NM3.LAST_NAME EMP_LNAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EN.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.LOCATION_SALE = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MF ON (EN.MASTER_FROMTO_ID = MF.MASTER_FROMTO_ID) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(FROM_COMMISSION_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(TO_COMMISSION_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
   ElseIf Ind = 4 Then                 'CReportCommission001
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME EMP_NAME, NM3.LAST_NAME EMP_LNAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EN.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MF ON (EN.MASTER_FROMTO_ID = MF.MASTER_FROMTO_ID) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(FROM_COMMISSION_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(TO_COMMISSION_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If MASTER_FROMTO_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") " & vbCrLf
   End If
   
   If COMMISSION_CHART_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.COMMISSION_CHART_ID = " & COMMISSION_CHART_ID & ") " & vbCrLf
   End If
   
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY EN.COMMISSION_CHART_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EN.COMMISSION_CHART_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY EN.COMMISSION_CHART_ID "
   ElseIf Ind = 4 Then
      'OrderByStr = "ORDER BY EN.COMMISSION_CHART_ID "
      OrderByStr = "ORDER BY EN.ORDER_ID "
   Else
      OrderByStr = "ORDER BY COMMISSION_CHART_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("COMMISSION_CHART_ID", NVLI(Rs("COMMISSION_CHART_ID"), -1))
      
      Call SetFieldValue("OLD_PK", GetFieldValue("COMMISSION_CHART_ID"))
      Call SetFieldValue("ORDER_ID", NVLI(Rs("ORDER_ID"), -1))
      Call SetFieldValue("MASTER_FROMTO_ID", NVLI(Rs("MASTER_FROMTO_ID"), -1))
      Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("CHILD_COUNT", NVLI(Rs("CHILD_COUNT"), -1))
      
      Call SetFieldValue("SALE_ID", NVLI(Rs("SALE_ID"), -1))
      Call SetFieldValue("RETURN_ID", NVLI(Rs("RETURN_ID"), -1))
      
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LNAME", NVLS(Rs("EMP_LNAME"), ""))
         
      Call SetFieldValue("SALE_NAME", NVLS(Rs("SALE_NAME"), ""))
      Call SetFieldValue("RETURN_NAME", NVLS(Rs("RETURN_NAME"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("COMMISSION_CHART_ID", NVLI(Rs("COMMISSION_CHART_ID"), -1))
      Call SetFieldValue("MASTER_FROMTO_ID", NVLI(Rs("MASTER_FROMTO_ID"), -1))
      Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("LOCATION_SALE", NVLI(Rs("LOCATION_SALE"), -1))
      Call SetFieldValue("SALE_TYPE", NVLI(Rs("SALE_TYPE"), -1))
      Call SetFieldValue("CHILD_COUNT", NVLI(Rs("CHILD_COUNT"), -1))
      
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LNAME", NVLS(Rs("EMP_LNAME"), ""))
         
   ElseIf Ind = 3 Then
      Call SetFieldValue("COMMISSION_CHART_ID", NVLI(Rs("COMMISSION_CHART_ID"), -1))
      Call SetFieldValue("MASTER_FROMTO_ID", NVLI(Rs("MASTER_FROMTO_ID"), -1))
      Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("LOCATION_SALE", NVLI(Rs("LOCATION_SALE"), -1))
      Call SetFieldValue("SALE_TYPE", NVLI(Rs("SALE_TYPE"), -1))
      Call SetFieldValue("CHILD_COUNT", NVLI(Rs("CHILD_COUNT"), -1))
      
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LNAME", NVLS(Rs("EMP_LNAME"), ""))
   ElseIf Ind = 4 Then
      Call SetFieldValue("COMMISSION_CHART_ID", NVLI(Rs("COMMISSION_CHART_ID"), -1))
      Call SetFieldValue("MASTER_FROMTO_ID", NVLI(Rs("MASTER_FROMTO_ID"), -1))
      Call SetFieldValue("PARENT_ID", NVLI(Rs("PARENT_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("CHILD_COUNT", NVLI(Rs("CHILD_COUNT"), -1))
      Call SetFieldValue("SALE_ID", NVLI(Rs("SALE_ID"), -1))
      Call SetFieldValue("RETURN_ID", NVLI(Rs("RETURN_ID"), -1))
      
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LNAME", NVLS(Rs("EMP_LNAME"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "COMMISSION_CHART"
   SequenceName = "COMMISSION_CHART_SEQ"
   PKField = "COMMISSION_CHART_ID"
   
   Call AddField("COMMISSION_CHART_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("ORDER_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("MASTER_FROMTO_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("PARENT_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("EMP_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("CHILD_COUNT", INT_TYPE, "", DATA_CAT)
   Call AddField("SALE_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("RETURN_ID", INT_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
   Call AddField("EMP_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_LNAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("OLD_PK", INT_TYPE, "", TEMP_CAT)
   
   Call AddField("FROM_COMMISSION_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_COMMISSION_DATE", DATE_TYPE, -1, TEMP_CAT)
   
   Call AddField("SALE_NAME", STRING_TYPE, "", TEMP_CAT)          'ประเภทชื่อการขาย
   Call AddField("RETURN_NAME", STRING_TYPE, "", TEMP_CAT)          'ประเภทชื่อการรับคืน
End Sub
