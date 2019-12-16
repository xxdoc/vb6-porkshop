VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCashDoc"
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

Public CashTranItems As Collection
Public TransferItems As Collection
Public JournalItems As Collection
Public PostItems As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set CashTranItems = New Collection
   Set TransferItems = New Collection
   Set JournalItems = New Collection
   Set PostItems = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set CashTranItems = Nothing
   Set TransferItems = Nothing
   Set JournalItems = Nothing
   Set PostItems = Nothing
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
Dim TempRs As ADODB.Recordset
Dim Ct As CCashTran
Dim iCount As Long
Dim Cq As CCheque

   Set Tf = m_FieldList(PKField)
   
   Set TempRs = New ADODB.Recordset
   Set Ct = New CCashTran
   Call Ct.SetFieldValue("CASH_DOC_ID", Tf.GetValue)
   Call Ct.QueryData(1, TempRs, iCount)
   Set Ct = Nothing
   While Not TempRs.EOF
      Set Ct = New CCashTran
      Call Ct.PopulateFromRS(1, TempRs)
      Call Ct.DeleteData
      If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
         Set Cq = New CCheque
         Call Cq.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
         Call Cq.SetFieldValue("BANK_FLAG", "N")
         Call Cq.SetFieldValue("POST_FLAG", "N")
         Call Cq.UpdateBankFlag
         Set Cq = Nothing
      End If
      Set Ct = Nothing
      Call TempRs.MoveNext
   Wend
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   SQL = "UPDATE CHEQUE CH SET CH.POST_FLAG = 'N',CH.BANK_FLAG = 'Y'  WHERE CH.CHEQUE_ID IN (SELECT CP.CHEQUE_ID FROM CASH_DOC_POST CP WHERE CP.CASH_DOC_ID = " & Tf.GetValue & ")"
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM CASH_DOC_POST  WHERE " & PKField & " = " & Tf.GetValue
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

Dim DOCUMENT_NO As String
Dim CASH_DOC_ID As Long
Dim DOCUMENT_TYPE As Long
Dim ENP_ID As Long
Dim ACCOUNT_NAME As String
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim TempDate As String

   CASH_DOC_ID = Val(GetFieldValue("CASH_DOC_ID"))
   DOCUMENT_TYPE = Val(GetFieldValue("DOCUMENT_TYPE"))
   DOCUMENT_NO = GetFieldValue("DOCUMENT_NO")
   ACCOUNT_NAME = GetFieldValue("ACCOUNT_NAME")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.KEY_NAME BANK_NAME, MR2.KEY_NAME BRANCH_NAME, MR3.KEY_NAME ACCOUNT_NAME, NM1.LONG_NAME EMP_NAME, NM1.LAST_NAME EMP_LASTNAME "
      SelectStr1 = SelectStr1 & ",NM2.LONG_NAME APAR_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (NM.BANK_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.BANK_BRANCH = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.BANK_ACCOUNT = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (EN.NAME_ID = NM1.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME APN ON (APN.APAR_MAS_ID = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (APN.NAME_ID = NM2.NAME_ID) "
      
      WhereStr = ""
   End If
   
   If CASH_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_ID = " & CASH_DOC_ID & ") " & vbCrLf
   End If
  
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
      
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(ACCOUNT_NAME)) > 0 Then
      If WildCard(ACCOUNT_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(MR3.KEY_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(MR3.KEY_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (MR3.KEY_NAME = '" & ChangeQuote(Trim(ACCOUNT_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (MR3.KEY_NAME = '" & ChangeQuote(Trim(ACCOUNT_NAME)) & "')"
         End If
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY NM.DOCUMENT_DATE ASC, NM.CASH_DOC_ID "
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

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CASH_DOC_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("DOCUMENT_DATE", InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), "")))
      Call SetFieldValue("DOCUMENT_NO", NVLS(Rs("DOCUMENT_NO"), ""))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("BANK_ACCOUNT", NVLI(Rs("BANK_ACCOUNT"), -1))
      Call SetFieldValue("DOCUMENT_TYPE", NVLI(Rs("DOCUMENT_TYPE"), -1))
      
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      
      Call SetFieldValue("KEY_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("DOCUMENT_NO"), ""))
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("ACCOUNT_NAME", NVLS(Rs("ACCOUNT_NAME"), ""))
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LASTNAME", NVLS(Rs("EMP_LASTNAME"), ""))
      Call SetFieldValue("EMP_NAME_LASTNAME", NVLS(Rs("EMP_NAME"), "") & " " & NVLS(Rs("EMP_LASTNAME"), ""))
      
      Call SetFieldValue("APAR_NAME", NVLS(Rs("APAR_NAME"), ""))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CASH_DOC"
   SequenceName = "CASH_DOC_SEQ"
   PKField = "CASH_DOC_ID"
   
   Call AddField("CASH_DOC_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("DOCUMENT_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("DOCUMENT_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("BANK_BRANCH", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ACCOUNT", INT_TYPE, -1, DATA_CAT)
   Call AddField("DOCUMENT_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("APAR_MAS_ID", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("ACCOUNT_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BANK_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_LASTNAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_NAME_LASTNAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("APAR_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CONFIG_DOC_TYPE", INT_TYPE, -1, TEMP_CAT)
   Call AddField("RUNNING_NO", INT_TYPE, "", TEMP_CAT)
   
End Sub
