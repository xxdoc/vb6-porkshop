VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRcpCnDn_Item"
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
Public CheckBillsFlag As Boolean

Public RCPCNDN_ITEM_ID As Long
Public BILLING_DOC_ID As Long
Public PAID_AMOUNT As Double
Public ITEM_AMOUNT As Double
Public TOTAL_PRICE As Double
Public VAT_AMOUNT As Double
Public DISCOUNT_AMOUNT As Double
Public EXT_DISCOUNT_AMOUNT As Double
Public SALE_CODE As String
Public CUSTOMER_BRANCH As Long
Public APAR_CODE As String
Public Due_Date As Date
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String

Public DOC_ID As Long

Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public DOCUMENT_TYPE As Long
Public DOCUMENT_TYPE_SET As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_APAR_CODE As String
Public TO_APAR_CODE As String
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

Dim RCPCNDN_ITEM_ID As Long
Dim BILLING_DOC_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim DOC_ID As Long
Dim DOCUMENT_TYPE As Long
Dim DOCUMENT_TYPE_SET As String
Dim FROM_DATE As Date
Dim TO_DATE As Date

   RCPCNDN_ITEM_ID = Val(GetFieldValue("RCPCNDN_ITEM_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   DOC_ID = Val(GetFieldValue("DOC_ID"))
   DOCUMENT_TYPE = Val(GetFieldValue("DOCUMENT_TYPE"))
   DOCUMENT_TYPE_SET = GetFieldValue("DOCUMENT_TYPE_SET")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   
   If Ind = 1 Then   ' NORMAL
      SelectStr1 = "SELECT EN.*,BD.DOCUMENT_NO DOC_NO,BD.DOCUMENT_DATE DOC_DATE,BD.DUE_DATE,BD.TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",BD.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",BD.DISCOUNT_AMOUNT,BD.EXT_DISCOUNT_AMOUNT,BD.VAT_AMOUNT,BD.PAY_AMOUNT "
      
      SelectStr1 = SelectStr1 & ",MR.KEY_NAME CNDN_REASON_NAME "
      SelectStr1 = SelectStr1 & ",BD.DOCUMENT_TYPE  DOC_ID_TYPE "
      SelectStr1 = SelectStr1 & ",BD.APAR_MAS_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (EN.DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (EN.CNDN_REASON = MR.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD1 ON (EN.BILLING_DOC_ID = BD1.BILLING_DOC_ID) "
      
      WhereStr = ""
      
      ' เป็นการเลือกเฉพาะ ใบส่งของจากใบวางบิลที่ใช้แล้วบางส่วน
      If CheckBillsFlag Then
         WhereStr = "WHERE ((SELECT COUNT(*) FROM RCPCNDN_ITEM RCP  LEFT OUTER JOIN BILLING_DOC BD10 ON (RCP.BILLING_DOC_ID = BD10.BILLING_DOC_ID) "
         WhereStr = WhereStr & " WHERE RCP.DOC_ID = EN.DOC_ID AND BD10.DOCUMENT_TYPE = " & RECEIPT2_DOCTYPE & " ) = 0 ) "
      End If
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EN.DOC_ID,MAX(EN.ITEM_AMOUNT) ITEM_AMOUNT , SUM(EN.PAID_DISCOUNT) PAID_DISCOUNT,SUM(EN.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
   End If
   
   If RCPCNDN_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.RCPCNDN_ITEM_ID = " & RCPCNDN_ITEM_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.DOC_ID = " & DOC_ID & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If Len(Trim(DOCUMENT_TYPE_SET)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_TYPE IN " & ChangeQuote(Trim(DOCUMENT_TYPE_SET)) & ")"
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY EN.DOC_ID "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY BD.APAR_MAS_ID,BD.DOCUMENT_NO "   'ใช้กับ  frmCreateBillingDocPack ต้องให้เรียงตาม APAR_MAS_ID เท่านั้น
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EN.DOC_ID "
   Else
      OrderByStr = "ORDER BY BD.APAR_MAS_ID,EN.RCPCNDN_ITEM_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, GetFieldValue("COUNT_FLAG"), ItemCount, Rs, glbErrorLog)
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
Dim TempDate As String
   
   
   If Ind = 3 Then     'GetPaidAmountByDocID
      SelectStr1 = "SELECT EN.DOC_ID,SUM(EN.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",BD2.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD1 ON (EN.BILLING_DOC_ID = BD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (BD1.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD2 ON (EN.DOC_ID = BD2.BILLING_DOC_ID) "
   ElseIf Ind = 4 Then     'Patch_2007_12_28_2_jill
      SelectStr1 = "SELECT EN.RCPCNDN_ITEM_ID,EN.PAID_AMOUNT,EN.ITEM_AMOUNT,BD2.TOTAL_PRICE,BD2.DISCOUNT_AMOUNT,BD2.EXT_DISCOUNT_AMOUNT,BD2.VAT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD2 ON (EN.DOC_ID = BD2.BILLING_DOC_ID) "
   ElseIf Ind = 5 Then     'GetPaidAmountBySaleCode
      SelectStr1 = "SELECT EN.DOC_ID,SUM(EN.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",BD2.DOCUMENT_TYPE,APM.APAR_CODE  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EN "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD1 ON (EN.BILLING_DOC_ID = BD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (BD1.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD2 ON (EN.DOC_ID = BD2.BILLING_DOC_ID) "
   End If

   If RCPCNDN_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.RCPCNDN_ITEM_ID = " & RCPCNDN_ITEM_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EN.DOC_ID = " & DOC_ID & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If Len(Trim(DOCUMENT_TYPE_SET)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_TYPE IN " & ChangeQuote(Trim(DOCUMENT_TYPE_SET)) & ")"
   End If
   
   If Len(Trim(FROM_APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_CODE >= '" & ChangeQuote(Trim(FROM_APAR_CODE)) & "')"
   End If
   
   If Len(Trim(TO_APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_CODE <= '" & ChangeQuote(Trim(TO_APAR_CODE)) & "')"
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BD1.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = "GROUP BY EN.DOC_ID "
      GroupByStr = GroupByStr & ",BD2.DOCUMENT_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY EN.DOC_ID,BD2.DOCUMENT_TYPE,APM.APAR_CODE "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY BD.APAR_MAS_ID,EN.RCPCNDN_ITEM_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EN.DOC_ID "
   Else
      OrderByStr = "ORDER BY BD.APAR_MAS_ID,EN.RCPCNDN_ITEM_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   If Ind = 3 Or Ind = 4 Or Ind = 5 Then
      OrderByStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("RCPCNDN_ITEM_ID", NVLI(Rs("RCPCNDN_ITEM_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("PAID_AMOUNT", NVLD(Rs("PAID_AMOUNT"), 0))
      Call SetFieldValue("CNDN_AMOUNT", NVLD(Rs("CNDN_AMOUNT"), 0))
      Call SetFieldValue("PAID_DISCOUNT", NVLD(Rs("PAID_DISCOUNT"), 0))
      Call SetFieldValue("PAID_DISCOUNT_PERCENT", NVLD(Rs("PAID_DISCOUNT_PERCENT"), 0))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("DOC_ID", NVLI(Rs("DOC_ID"), -1))
      Call SetFieldValue("CNDN_REASON", NVLI(Rs("CNDN_REASON"), -1))
      Call SetFieldValue("BILLS_ID", NVLI(Rs("BILLS_ID"), -1))
      Call SetFieldValue("BILLS_NO", NVLS(Rs("BILLS_NO"), ""))
      Call SetFieldValue("DOC_ID_RCP", NVLI(Rs("DOC_ID_RCP"), -1))
      Call SetFieldValue("DOC_ID_BILLS", NVLI(Rs("DOC_ID_BILLS"), -1))
      
      Call SetFieldValue("DOC_NO", NVLS(Rs("DOC_NO"), ""))
      Call SetFieldValue("DOC_DATE", InternalDateToDate(NVLS(Rs("DOC_DATE"), "")))
      Call SetFieldValue("CNDN_REASON_NAME", NVLS(Rs("CNDN_REASON_NAME"), ""))
      Call SetFieldValue("DUE_DATE", InternalDateToDate(NVLS(Rs("DUE_DATE"), "")))
      
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("DISCOUNT_AMOUNT", NVLD(Rs("DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("EXT_DISCOUNT_AMOUNT", NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0))
      Call SetFieldValue("VAT_AMOUNT", NVLD(Rs("VAT_AMOUNT"), 0))
      Call SetFieldValue("PAY_AMOUNT", NVLD(Rs("PAY_AMOUNT"), 0))
      
      Call SetFieldValue("DOC_ID_TYPE", NVLI(Rs("DOC_ID_TYPE"), -1))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      
      Call SetFieldValue("DOCUMENT_TYPE", NVLI(Rs("DOCUMENT_TYPE"), -1))
      
   ElseIf Ind = 2 Then
      Call SetFieldValue("PAID_DISCOUNT", NVLD(Rs("PAID_DISCOUNT"), 0))
      Call SetFieldValue("PAID_AMOUNT", NVLD(Rs("PAID_AMOUNT"), 0))
      Call SetFieldValue("DOC_ID", NVLI(Rs("DOC_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
   ElseIf Ind = 3 Then
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DOC_ID = NVLI(Rs("DOC_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 4 Then
      RCPCNDN_ITEM_ID = NVLI(Rs("RCPCNDN_ITEM_ID"), -1)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXT_DISCOUNT_AMOUNT = NVLD(Rs("EXT_DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DOC_ID = NVLI(Rs("DOC_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
   End If
End Sub

Public Sub RegisterFields()
   TableName = "RCPCNDN_ITEM"
   SequenceName = "RCPCNDN_ITEM_SEQ"
   PKField = "RCPCNDN_ITEM_ID"
   
   'If MasterInd = 1 Or MasterInd = 2 Then
      Call AddField("RCPCNDN_ITEM_ID", INT_TYPE, -1, ID_CAT)
      Call AddField("BILLING_DOC_ID", INT_TYPE, "", DATA_CAT)
      Call AddField("PAID_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("CNDN_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("PAID_DISCOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("PAID_DISCOUNT_PERCENT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("DOC_ID", INT_TYPE, -1, DATA_CAT)          'อ้างอิงเอกสาร
      Call AddField("CNDN_REASON", INT_TYPE, -1, DATA_CAT)
      Call AddField("BILLS_ID", INT_TYPE, -1, DATA_CAT)        'รับชำระจากใบวางบิลใบไหน
      Call AddField("BILLS_NO", STRING_TYPE, "", DATA_CAT)     'รับชำระจากใบวางบิลใบไหน
      Call AddField("DOC_ID_BILLS", INT_TYPE, -1, DATA_CAT)       'แยกให้รู้ว่าอ้างอิงของใบวางบิลใบไหน
      Call AddField("DOC_ID_RCP", INT_TYPE, -1, DATA_CAT)            'ให้รู้ว่าอ้างอิงใบเสร็จใบไหน
      Call AddField("BILLING_DOC_PACK", INT_TYPE, -1, DATA_CAT)         'บอกให้รู้ว่าสร้างมาจากเอกสารเป็นชุด
      
      Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
      Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
      Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
      Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
      
      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
      Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
      
      Call AddField("DOC_NO", STRING_TYPE, "", TEMP_CAT)
      Call AddField("DOC_DATE", DATE_TYPE, -1, TEMP_CAT)
      
      Call AddField("CNDN_REASON_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("DOCUMENT_TYPE", INT_TYPE, -1, TEMP_CAT)
      Call AddField("DOC_ID_TYPE", INT_TYPE, -1, TEMP_CAT)
      Call AddField("DUE_DATE", DATE_TYPE, -1, TEMP_CAT)
      Call AddField("TOTAL_PRICE", MONEY_TYPE, 0, TEMP_CAT)
      Call AddField("DISCOUNT_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      Call AddField("EXT_DISCOUNT_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      Call AddField("VAT_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      Call AddField("PAY_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      
      Call AddField("DOCUMENT_TYPE_SET", STRING_TYPE, "", TEMP_CAT)
      
      Call AddField("SELECT_FLAG", STRING_TYPE, "Y", TEMP_CAT)
      Call AddField("APAR_MAS_ID", INT_TYPE, -1, TEMP_CAT)
      
      ' สถานะเพื่อให้นับหรือ ไม่นับเพื่อความเร็ว
      Call AddField("COUNT_FLAG", BOOLEAN_TYPE, True, TEMP_CAT)
      ' สถานะเพื่อให้นับหรือ ไม่นับเพื่อความเร็ว
   
   'End If
End Sub

Public Function CopyObject(Ind As Long, TempObj As CRcpCnDn_Item)
      'Call SetFieldValue("RCPCNDN_ITEM_ID", TempObj.GetFieldValue("RCPCNDN_ITEM_ID")) ห้าม COPY ไปเพราะว่าจะมีผลกับการ ลบข้อมูล
      'Call SetFieldValue("BILLING_DOC_ID", TempObj.GetFieldValue("BILLING_DOC_ID")) ห้าม COPY ไปเพราะว่าจะมีผลกับการ ลบข้อมูล
      Call SetFieldValue("PAID_AMOUNT", TempObj.GetFieldValue("PAID_AMOUNT"))
      Call SetFieldValue("CNDN_AMOUNT", TempObj.GetFieldValue("CNDN_AMOUNT"))
      Call SetFieldValue("PAID_DISCOUNT", TempObj.GetFieldValue("PAID_DISCOUNT"))
      Call SetFieldValue("PAID_DISCOUNT_PERCENT", TempObj.GetFieldValue("PAID_DISCOUNT_PERCENT"))
      Call SetFieldValue("ITEM_AMOUNT", TempObj.GetFieldValue("ITEM_AMOUNT"))
      Call SetFieldValue("DOC_ID", TempObj.GetFieldValue("DOC_ID"))
      Call SetFieldValue("CNDN_REASON", TempObj.GetFieldValue("CNDN_REASON"))
      
      Call SetFieldValue("DOC_NO", TempObj.GetFieldValue("DOC_NO"))
      Call SetFieldValue("DOC_DATE", TempObj.GetFieldValue("DOC_DATE"))
      
      Call SetFieldValue("DOC_ID_TYPE", TempObj.GetFieldValue("DOC_ID_TYPE"))
End Function
Public Function UpdateDocIDRcp() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE RCPCNDN_ITEM SET " & _
                  "   DOC_ID_RCP = " & GetFieldValue("DOC_ID_RCP") & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE RCPCNDN_ITEM_ID = " & GetFieldValue("RCPCNDN_ITEM_ID")

   Call m_ADOConn.Execute(SQL1)

End Function
Public Function UpdatePaidItemAmount() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE RCPCNDN_ITEM SET " & _
                  "   PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                  "   ITEM_AMOUNT = " & ITEM_AMOUNT & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE RCPCNDN_ITEM_ID = " & RCPCNDN_ITEM_ID

   Call m_ADOConn.Execute(SQL1)

End Function
Public Function DeleteByDocID() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
   
   SQL1 = "DELETE FROM RCPCNDN_ITEM WHERE DOC_ID = " & DOC_ID
   Call m_ADOConn.Execute(SQL1)

End Function
