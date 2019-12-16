VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillingSubTract"
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

Public BILLING_SUBTRACT_ID As Long
Public BILLING_DOC_ID As Long
Public SUBTRACT_ID As Long
Public ITEM_AMOUNT As Double
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public FROM_DATE As Date
Public TO_DATE As Date
Public SUBTRACT_CODE As String
Public SUBTRACT_NAME As String

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

Dim BILLING_SUBTRACT_ID As Long
Dim BILLING_DOC_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   BILLING_SUBTRACT_ID = Val(GetFieldValue("BILLING_SUBTRACT_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EN.* "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE SUBTRACT_CODE,MR.KEY_NAME SUBTRACT_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (EN.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.SUBTRACT_ID = MR.KEY_ID) "
      
      WhereStr = ""
      
   End If
   
   If BILLING_SUBTRACT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_SUBTRACT_ID = " & BILLING_SUBTRACT_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   
   GroupByStr = ""
   
   If Ind = 1 Then
      'OrderByStr = "ORDER BY BILLING_SUBTRACT_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      'OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, False, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub QueryDataReport(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional CountFlag As Boolean = False)
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
   
   If Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT EN.SUBTRACT_ID "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE SUBTRACT_CODE,MR.KEY_NAME SUBTRACT_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (EN.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.SUBTRACT_ID = MR.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT SUM(EN.ITEM_AMOUNT) ITEM_AMOUNT,EN.SUBTRACT_ID "
      SelectStr1 = SelectStr1 & ",BD.BILLING_DOC_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (EN.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   End If
   
   If BILLING_SUBTRACT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_SUBTRACT_ID = " & BILLING_SUBTRACT_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = " GROUP BY EN.SUBTRACT_ID,BD.BILLING_DOC_ID "
   End If
   
   If Ind = 2 Then
      OrderByStr = "ORDER BY MR.KEY_CODE "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   If Ind = 3 Then
      OrderByStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, CountFlag, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("BILLING_SUBTRACT_ID", NVLI(Rs("BILLING_SUBTRACT_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("SUBTRACT_ID", NVLI(Rs("SUBTRACT_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      
      Call SetFieldValue("SUBTRACT_CODE", NVLS(Rs("SUBTRACT_CODE"), ""))
      Call SetFieldValue("SUBTRACT_NAME", NVLS(Rs("SUBTRACT_NAME"), ""))
   ElseIf Ind = 2 Then
      SUBTRACT_ID = NVLI(Rs("SUBTRACT_ID"), -1)
      SUBTRACT_CODE = NVLS(Rs("SUBTRACT_CODE"), "")
      SUBTRACT_NAME = NVLS(Rs("SUBTRACT_NAME"), "")
   ElseIf Ind = 3 Then
      SUBTRACT_ID = NVLI(Rs("SUBTRACT_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
   End If
End Sub

Public Sub RegisterFields()
   TableName = "BILLING_SUBTRACT"
   SequenceName = "BILLING_SUBTRACT_SEQ"
   PKField = "BILLING_SUBTRACT_ID"
   
   Call AddField("BILLING_SUBTRACT_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BILLING_DOC_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("SUBTRACT_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BILLING_DOC_PACK", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
   Call AddField("SUBTRACT_CODE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("SUBTRACT_NAME", INT_TYPE, -1, TEMP_CAT)
   
End Sub
