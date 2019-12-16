VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheque"
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

Public CHEQUE_NO As String
Public CHEQUE_ID As Long
Public CHEQUE_TYPE As Long
Public CHEQUE_AMOUNT As Double
Public CHEQUE_DATE As Date
Public Direction As Long
Public BANK_FLAG As String
Public POST_FLAG As String
Public APAR_MAS_ID As Long
Public EFFECTIVE_DATE As Date

'---------------------------------------------------------------------------------------------------------------------------------------------------------->
Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DATE2 As Date
Public TO_DATE2 As Date
Public APAR_CODE As String
Public APAR_NAME As String
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public APAR_TYPE As Long
Public TEMP_FEE_AMOUNT As Double


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
Dim BANK_FLAG As String
Dim POST_FLAG As String

   Set Tf = m_FieldList(PKField)
   
   Call PopulateInternalField(ShowMode, Me)
   
   BANK_FLAG = GetFieldValue("BANK_FLAG")
   If BANK_FLAG = "" Then
      BANK_FLAG = "N"
      Call SetFieldValue("BANK_FLAG", BANK_FLAG)
   End If
   
   POST_FLAG = GetFieldValue("POST_FLAG")
   If POST_FLAG = "" Then
      POST_FLAG = "N"
      Call SetFieldValue("POST_FLAG", POST_FLAG)
   End If
   
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

Dim CHEQUE_NO As String
Dim CHEQUE_ID As Long
Dim CHEQUE_TYPE As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim FROM_DATE2 As Date
Dim TO_DATE2 As Date
Dim APAR_CODE As String
Dim APAR_NAME As String
Dim Direction As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim APAR_TYPE As Long
Dim BANK_FLAG As String
Dim POST_FLAG As String
Dim APAR_MAS_ID As Long
   
   CHEQUE_ID = Val(GetFieldValue("CHEQUE_ID"))
   CHEQUE_TYPE = Val(GetFieldValue("CHEQUE_TYPE"))
   CHEQUE_NO = GetFieldValue("CHEQUE_NO")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   FROM_DATE2 = GetFieldValue("FROM_DATE2")
   TO_DATE2 = GetFieldValue("TO_DATE2")
   APAR_NAME = GetFieldValue("APAR_NAME")
   APAR_CODE = GetFieldValue("APAR_CODE")
   APAR_TYPE = Val(GetFieldValue("APAR_TYPE"))
   BANK_FLAG = GetFieldValue("BANK_FLAG")
   POST_FLAG = GetFieldValue("POST_FLAG")
   Direction = Val(GetFieldValue("DIRECTION"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   APAR_MAS_ID = Val(GetFieldValue("APAR_MAS_ID"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, APM.APAR_NAME, APM.APAR_CODE ,MR1.KEY_NAME BANK_NAME, MR2.KEY_NAME BRANCH_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.BANK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.BANK_BRANCH = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT NM.*, APM.APAR_NAME, APM.APAR_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT NM.*, APM.APAR_NAME,APM.APAR_CODE, MR1.KEY_NAME BANK_NAME, MR2.KEY_NAME BRANCH_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.BANK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.BANK_BRANCH = MR2.KEY_ID) "
   End If
   
   If CHEQUE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_ID = " & CHEQUE_ID & ") " & vbCrLf
   End If
   
   If CHEQUE_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_TYPE = " & CHEQUE_TYPE & ") " & vbCrLf
   End If
   
   If APAR_MAS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_MAS_ID = " & APAR_MAS_ID & ") " & vbCrLf
   End If
   
   If APAR_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_TYPE = " & APAR_TYPE & ") " & vbCrLf
   End If
      
   If Len(Trim(CHEQUE_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CHEQUE_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.CHEQUE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_NO)) & "')"
      End If
   End If
      
   If Len(Trim(APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
      End If
   End If
   
   If Len(Trim(APAR_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APM.APAR_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APM.APAR_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If FROM_DATE2 > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE2)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EFFECTIVE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE2 > 0 Then
      TempDate = DateToStringIntHi(TO_DATE2)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EFFECTIVE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Direction > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DIRECTION = " & Direction & ") " & vbCrLf
   End If
   
   If Len(Trim(BANK_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(BANK_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.BANK_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.BANK_FLAG = '" & ChangeQuote(Trim(BANK_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(POST_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(BANK_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.CHEQUE_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY NM.CHEQUE_DATE "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY APM.APAR_CODE ASC, NM2.LONG_NAME ASC, NM.CHEQUE_DATE ASC, NM.CHEQUE_NO "
      Else
         OrderByStr = "ORDER BY NM.CHEQUE_NO "
      End If
   ElseIf Ind = 2 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE,NM.CHEQUE_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APM.APAR_CODE,NM.CHEQUE_DATE "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE,NM.CHEQUE_NO "
      End If
   ElseIf Ind = 3 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE ASC, APM.APAR_MAS_ID ASC, NM.CHEQUE_DATE ASC, NM.CHEQUE_NO "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE ASC, APM.APAR_MAS_ID ASC, NM.CHEQUE_DATE ASC, NM.CHEQUE_NO "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub QueryDataReport(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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

   
   If Ind = 4 Then
      SelectStr1 = "SELECT NM.CHEQUE_ID,NM.CHEQUE_NO,NM.CHEQUE_DATE,NM.CHEQUE_AMOUNT,NM.EFFECTIVE_DATE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
   
   If CHEQUE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_ID = " & CHEQUE_ID & ") " & vbCrLf
   End If
   
   If CHEQUE_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_TYPE = " & CHEQUE_TYPE & ") " & vbCrLf
   End If
   
   If APAR_MAS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.APAR_MAS_ID = " & APAR_MAS_ID & ") " & vbCrLf
   End If
   
   If APAR_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_TYPE = " & APAR_TYPE & ") " & vbCrLf
   End If
      
   If Len(Trim(CHEQUE_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CHEQUE_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.CHEQUE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_NO)) & "')"
      End If
   End If
      
   If Len(Trim(APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APM.APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APM.APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
      End If
   End If
   
   If Len(Trim(APAR_NAME)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_NAME, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APM.APAR_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APM.APAR_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CHEQUE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If FROM_DATE2 > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE2)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EFFECTIVE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE2 > 0 Then
      TempDate = DateToStringIntHi(TO_DATE2)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EFFECTIVE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Direction > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DIRECTION = " & Direction & ") " & vbCrLf
   End If
   
   If Len(Trim(BANK_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(BANK_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.BANK_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.BANK_FLAG = '" & ChangeQuote(Trim(BANK_FLAG)) & "')"
      End If
   End If
   
   If Len(Trim(POST_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(BANK_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(NM.POST_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (NM.POST_FLAG = '" & ChangeQuote(Trim(POST_FLAG)) & "')"
      End If
   End If
   
   If Ind = 4 Then
         OrderByStr = "ORDER BY NM.CHEQUE_DATE,NM.CHEQUE_NO "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CHEQUE_ID", NVLI(Rs("CHEQUE_ID"), -1))
      Call SetFieldValue("CHEQUE_NO", NVLS(Rs("CHEQUE_NO"), ""))
      Call SetFieldValue("CHEQUE_DATE", InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), "")))
      Call SetFieldValue("CHEQUE_TYPE", NVLI(Rs("CHEQUE_TYPE"), -1))
      Call SetFieldValue("EFFECTIVE_DATE", InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), "")))
      Call SetFieldValue("CHEQUE_STATUS", NVLI(Rs("CHEQUE_STATUS"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("DIRECTION", NVLI(Rs("DIRECTION"), -1))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("CHEQUE_AMOUNT", NVLD(Rs("CHEQUE_AMOUNT"), 0))
      Call SetFieldValue("BANK_FLAG", NVLS(Rs("BANK_FLAG"), ""))
      Call SetFieldValue("POST_FLAG", NVLS(Rs("POST_FLAG"), ""))

      Call SetFieldValue("APAR_NAME", NVLS(Rs("APAR_NAME"), ""))
      Call SetFieldValue("APAR_CODE", NVLS(Rs("APAR_CODE"), ""))
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
   ElseIf Ind = 3 Then
      Call SetFieldValue("CHEQUE_ID", NVLI(Rs("CHEQUE_ID"), -1))
      Call SetFieldValue("CHEQUE_NO", NVLS(Rs("CHEQUE_NO"), ""))
      Call SetFieldValue("CHEQUE_DATE", InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), "")))
      Call SetFieldValue("CHEQUE_TYPE", NVLI(Rs("CHEQUE_TYPE"), -1))
      Call SetFieldValue("EFFECTIVE_DATE", InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), "")))
      Call SetFieldValue("CHEQUE_STATUS", NVLI(Rs("CHEQUE_STATUS"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("DIRECTION", NVLI(Rs("DIRECTION"), -1))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("CHEQUE_AMOUNT", NVLD(Rs("CHEQUE_AMOUNT"), 0))
      Call SetFieldValue("BANK_FLAG", NVLS(Rs("BANK_FLAG"), ""))

      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("APAR_NAME", NVLS(Rs("APAR_NAME"), ""))
      Call SetFieldValue("APAR_CODE", NVLS(Rs("APAR_CODE"), ""))
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
    ElseIf Ind = 4 Then
      CHEQUE_ID = NVLI(Rs("CHEQUE_ID"), -1)
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      CHEQUE_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), ""))
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      EFFECTIVE_DATE = InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), ""))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CHEQUE"
   SequenceName = "CHEQUE_SEQ"
   PKField = "CHEQUE_ID"
   
      Call AddField("CHEQUE_ID", INT_TYPE, -1, ID_CAT)
      Call AddField("CHEQUE_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("CHEQUE_DATE", DATE_TYPE, -1, DATA_CAT)
      Call AddField("CHEQUE_TYPE", INT_TYPE, -1, DATA_CAT)
      Call AddField("EFFECTIVE_DATE", DATE_TYPE, -1, DATA_CAT)
      Call AddField("CHEQUE_STATUS", INT_TYPE, -1, DATA_CAT)
      Call AddField("BANK_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("BANK_BRANCH", INT_TYPE, -1, DATA_CAT)
      Call AddField("DIRECTION", INT_TYPE, -1, DATA_CAT)
      Call AddField("CHEQUE_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("APAR_MAS_ID", INT_TYPE, -1, DATA_CAT)
      Call AddField("BANK_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("POST_FLAG", STRING_TYPE, "", DATA_CAT)
      Call AddField("BILLING_DOC_TYPE", INT_TYPE, -1, DATA_CAT)
      Call AddField("BILLING_DOC_PACK", INT_TYPE, -1, DATA_CAT)
      
      Call AddField("TX_NO", STRING_TYPE, "", DATA_CAT)
      Call AddField("TX_DATE", DATE_TYPE, -1, DATA_CAT)
      
      Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
      Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
      Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
      Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
      Call AddField("APAR_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("APAR_CODE", STRING_TYPE, "", TEMP_CAT)
      Call AddField("APAR_TYPE", INT_TYPE, -1, TEMP_CAT)
      Call AddField("BANK_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("TEMP_FEE_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      
      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
      Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("FROM_DATE2", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TO_DATE2", DATE_TYPE, -1, TEMP_CAT)
End Sub
Public Sub UpdateApArID()
Dim SQL As String
Dim APAR_MAS_ID As Long
Dim CHEQUE_ID As Long

   APAR_MAS_ID = GetFieldValue("APAR_MAS_ID")
   CHEQUE_ID = GetFieldValue("CHEQUE_ID")
   If APAR_MAS_ID > 0 Then
      SQL = "UPDATE CHEQUE SET APAR_MAS_ID = " & APAR_MAS_ID & " WHERE CHEQUE_ID = " & CHEQUE_ID & ""
      Call m_ADOConn.Execute(SQL)
   End If
End Sub
Public Sub UpdateBankBranch()
Dim SQL As String
Dim BANK_ID As Long
Dim BANK_BRANCH As Long
Dim CHEQUE_ID As Long

   BANK_ID = GetFieldValue("BANK_ID")
   BANK_BRANCH = GetFieldValue("BANK_BRANCH")
   CHEQUE_ID = GetFieldValue("CHEQUE_ID")
   
   SQL = "UPDATE CHEQUE SET BANK_ID = " & BANK_ID & ", BANK_BRANCH = " & BANK_BRANCH & " WHERE CHEQUE_ID = " & CHEQUE_ID & ""
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub UpdateBankFlag()
Dim SQL As String
Dim BANK_FLAG As String
Dim CHEQUE_ID As Long
Dim POST_FLAG As String

   BANK_FLAG = GetFieldValue("BANK_FLAG")
   CHEQUE_ID = GetFieldValue("CHEQUE_ID")
   POST_FLAG = GetFieldValue("POST_FLAG")
   
   SQL = "UPDATE CHEQUE SET POST_FLAG = '" & POST_FLAG & "' WHERE CHEQUE_ID = " & CHEQUE_ID & ""
   Call m_ADOConn.Execute(SQL)
   
   SQL = "UPDATE CHEQUE SET BANK_FLAG = '" & BANK_FLAG & "' WHERE CHEQUE_ID = " & CHEQUE_ID & ""
   Call m_ADOConn.Execute(SQL)
   
   
End Sub


