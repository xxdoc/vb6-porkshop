VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE
Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection
Public TagetDetails As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set TagetDetails = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set TagetDetails = Nothing
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
'

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
   
   SQL = "DELETE FROM TAGET_DETAIL  WHERE " & PKField & " = " & Tf.GetValue
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

Dim TAGET_ID As Long
Dim EMP_ID As Long
Dim MONTH_ID As Long
Dim YEAR_NO As String
Dim EMP_CODE As String
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim EMP_NAME As String
Dim EMP_LNAME As String
Dim TAGET_TYPE As Long

   TAGET_ID = Val(GetFieldValue("TAGET_ID"))
   EMP_ID = Val(GetFieldValue("EMP_ID"))
   MONTH_ID = Val(GetFieldValue("MONTH_ID"))
   YEAR_NO = GetFieldValue("YEAR_NO")
   EMP_CODE = GetFieldValue("EMP_CODE")
   EMP_NAME = GetFieldValue("EMP_NAME")
   EMP_LNAME = GetFieldValue("EMP_LNAME")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   TAGET_TYPE = Val(GetFieldValue("TAGET_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, NM.LONG_NAME EMP_NAME, NM.LAST_NAME EMP_LNAME "
      SelectStr1 = SelectStr1 & ",EM1.EMP_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM1 ON (EM.EMP_ID = EM1.EMP_ID) "
      
      
      
      WhereStr = ""
   End If
   
   
   If TAGET_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.TAGET_ID = " & TAGET_ID & ") " & vbCrLf
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.EMP_ID = " & EMP_ID & ") " & vbCrLf
   End If
   
   If MONTH_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.MONTH_ID = " & MONTH_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(YEAR_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(YEAR_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(EM.YEAR_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "') "
      Else
         WhereStr = WhereStr & TempStr & " (EM.YEAR_NO = '" & ChangeQuote(Trim(YEAR_NO)) & "') "
      End If
   End If
   
   If Len(Trim(EMP_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(EM1.EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (EM1.EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
      End If
   End If
   
   If Len(Trim(EMP_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(EMP_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.LONG_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
      End If
   End If
   
   If Len(Trim(EMP_LNAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(EMP_LNAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.LAST_NAME = '" & ChangeQuote(Trim(EMP_LNAME)) & "')"
      End If
   End If
   
   If TAGET_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.TAGET_TYPE = " & TAGET_TYPE & ") " & vbCrLf
   End If
   
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EM.YYYYMM "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM1.EMP_CODE "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY NM.LONG_NAME "
      Else
         OrderByStr = "ORDER BY EM.TAGET_ID "
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
      Call SetFieldValue("TAGET_ID", NVLI(Rs("TAGET_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("MONTH_ID", NVLI(Rs("MONTH_ID"), -1))
      
      Call SetFieldValue("YEAR_NO", NVLS(Rs("YEAR_NO"), ""))
      Call SetFieldValue("YYYYMM", NVLS(Rs("YYYYMM"), ""))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TAGET_DESC", NVLS(Rs("TAGET_DESC"), ""))
      
      Call SetFieldValue("TAGET_TYPE", NVLI(Rs("TAGET_TYPE"), -1))
      
      Call SetFieldValue("EMP_CODE", NVLS(Rs("EMP_CODE"), ""))
      
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LNAME", NVLS(Rs("EMP_LNAME"), ""))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "TAGET"
   SequenceName = "TAGET_SEQ"
   PKField = "TAGET_ID"
   
   Call AddField("TAGET_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("EMP_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("MONTH_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("YEAR_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("YYYYMM", STRING_TYPE, "", DATA_CAT)
   Call AddField("TAGET_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("TAGET_DESC", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("TOTAL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("EMP_CODE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_LNAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
End Sub
