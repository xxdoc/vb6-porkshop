VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTagetDetail"
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

Public BRANCH_CODE As String
Public APAR_CODE As String
Public APAR_GROUP_NAME As String
Public EMPLOYEE_CODE As String
Public STOCK_NO As String
Public STOCK_DESC As String
Public JOINT_CODE As String
Public UNIT_AMOUNT As String
Public BILL_DESC As String
Public STOCK_TYPE_CODE As String
Public STOCK_GROUP_NAME As String

Public TOTAL_AMOUNT As Double
Public TOTAL_PRICE As Double
Public TOTAL_AMOUNT_RT As Double
Public TOTAL_PRICE_RT As Double
Public SALE_JOINT_CODE As String
Public SALE_LONG_NAME As String
Public SALE_LAST_NAME As String

Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String

Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String

Public TOP_FLAG As String           'สำหรับลูกค้า TOP
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
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional CountFlag As Boolean = True)
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

Dim TAGET_DETAIL_ID As Long
Dim TAGET_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim YYYYMM As String
   
   TAGET_DETAIL_ID = Val(GetFieldValue("TAGET_DETAIL_ID"))
   TAGET_ID = Val(GetFieldValue("TAGET_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   YYYYMM = GetFieldValue("YYYYMM")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE,MR.KEY_NAME BRANCH_NAME "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE APAR_CODE,NM2.LONG_NAME APAR_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE,NM3.LONG_NAME EMPLOYEE_LONG_NAME,NM3.LAST_NAME EMPLOYEE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",NM.EMP_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE,MR.KEY_NAME BRANCH_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE,NM3.LONG_NAME EMPLOYEE_LONG_NAME,NM3.LAST_NAME EMPLOYEE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
      
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.STOCK_DESC,SC.BILL_DESC "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 9 Then  'LoadTagetDetailBranchStock
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",NM.EMP_ID,SC.GROUP_COM_ID "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,MR3.KEY_NAME APAR_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (MR2.PARENT_ID = MR3.KEY_ID) "
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",MR6.KEY_CODE STOCK_TYPE_CODE,MR6.KEY_NAME STOCK_GROUP_NAME  "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (SC.STOCK_TYPE = MR5.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR6 ON (MR5.PARENT_ID = MR6.KEY_ID) "
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT SUM(NM.TOTAL_AMOUNT) TOTAL_AMOUNT,SUM(NM.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr1 = SelectStr1 & ",SUM(NM.TOTAL_AMOUNT_RT) TOTAL_AMOUNT_RT,SUM(NM.TOTAL_PRICE_RT) TOTAL_PRICE_RT "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE EMPLOYEE_CODE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 17 Then
      SelectStr1 = "SELECT EM.EMP_CODE EMPLOYEE_CODE,EM.JOINT_CODE SALE_JOINT_CODE "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO,SC.JOINT_CODE,SC.STOCK_DESC,SC.UNIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID ) "
   ElseIf Ind = 18 Then
      SelectStr1 = "SELECT EM.EMP_CODE EMPLOYEE_CODE,EM.JOINT_CODE SALE_JOINT_CODE "
      SelectStr1 = SelectStr1 & ",NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET TG ON (NM.TAGET_ID = TG.TAGET_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.KEY_ID = NM.BRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.APAR_MAS_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
   End If
   
   If TAGET_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.TAGET_DETAIL_ID = " & TAGET_DETAIL_ID & ") " & vbCrLf
   End If
   
   If TAGET_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.TAGET_ID = " & TAGET_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(YYYYMM)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(YYYYMM, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(TG.YYYYMM,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "') "
      Else
         WhereStr = WhereStr & TempStr & " (TG.YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "') "
      End If
   End If
   
   If Len(Trim(TOP_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      Call AddSearchAparTop(FromStr)
      WhereStr = WhereStr & TempStr & " (APM.LABEL_FLAG = '" & ChangeQuote(Trim(TOP_FLAG)) & "') "
   End If
   
   If Len(Trim(FROM_SALE_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.EMP_CODE >= '" & ChangeQuote(Trim(FROM_SALE_CODE)) & "')"
   End If
   
   If Len(Trim(TO_SALE_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.EMP_CODE <= '" & ChangeQuote(Trim(TO_SALE_CODE)) & "')"
   End If
   
   If Trim(FROM_STOCK_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(TO_STOCK_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      End If
   End If
   
   If Ind = 2 Then
      GroupByStr = " GROUP BY NM.EMP_ID "
   ElseIf Ind = 3 Then
      GroupByStr = " GROUP BY MR.KEY_CODE "
      GroupByStr = GroupByStr & ",APM.APAR_CODE "
      GroupByStr = GroupByStr & ",EM.EMP_CODE  "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 4 Then
      GroupByStr = " GROUP BY MR.KEY_CODE,MR.KEY_NAME "
      GroupByStr = GroupByStr & ",EM.EMP_CODE ,NM3.LONG_NAME ,NM3.LAST_NAME "
      GroupByStr = GroupByStr & ",SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
   ElseIf Ind = 5 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,SC.UNIT_AMOUNT "
   ElseIf Ind = 6 Then
      GroupByStr = " GROUP BY MR.KEY_CODE "
      GroupByStr = GroupByStr & ",EM.EMP_CODE  "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 7 Then
      GroupByStr = " GROUP BY SC.STOCK_NO,SC.STOCK_DESC,SC.BILL_DESC "
   ElseIf Ind = 8 Then
      GroupByStr = " GROUP BY MR.KEY_CODE "
      GroupByStr = GroupByStr & ",APM.APAR_CODE "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 9 Then
      GroupByStr = " GROUP BY MR.KEY_CODE "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 10 Then
      GroupByStr = " GROUP BY APM.APAR_CODE "
      GroupByStr = GroupByStr & ",EM.EMP_CODE  "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 11 Then
      GroupByStr = " GROUP BY EM.EMP_CODE "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 12 Then
      GroupByStr = " GROUP BY APM.APAR_CODE "
      GroupByStr = GroupByStr & ",SC.STOCK_NO "
   ElseIf Ind = 13 Then
      GroupByStr = " GROUP BY NM.EMP_ID,SC.GROUP_COM_ID "
   ElseIf Ind = 14 Then
      GroupByStr = " GROUP BY MR3.KEY_NAME, EM.EMP_CODE, SC.STOCK_NO "
   ElseIf Ind = 15 Then
      GroupByStr = " GROUP BY MR6.KEY_CODE,MR6.KEY_NAME "
   ElseIf Ind = 16 Then
      GroupByStr = " GROUP BY EM.EMP_CODE "
   ElseIf Ind = 18 Then
      GroupByStr = " GROUP BY EM.EMP_CODE,EM.JOINT_CODE,NM3.LONG_NAME,NM3.LAST_NAME "
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY NM.TAGET_DETAIL_ID "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY NM.EMP_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY MR.KEY_CODE,EM.EMP_CODE,SC.STOCK_NO "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY EM.EMP_CODE,SC.STOCK_NO "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   If Ind = 3 Or Ind = 6 Or Ind = 7 Or Ind = 8 Or Ind = 9 Or Ind = 10 Or Ind = 11 Or Ind = 12 Or Ind = 13 Or Ind = 14 Or Ind = 15 Or Ind = 16 Or Ind = 18 Then
      OrderByStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, CountFlag, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("TAGET_DETAIL_ID", NVLI(Rs("TAGET_DETAIL_ID"), -1))
      Call SetFieldValue("TAGET_ID", NVLI(Rs("TAGET_ID"), -1))
      Call SetFieldValue("BRANCH_ID", NVLI(Rs("BRANCH_ID"), -1))
      Call SetFieldValue("APAR_MAS_ID", NVLI(Rs("APAR_MAS_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT_RT", NVLD(Rs("TOTAL_AMOUNT_RT"), 0))
      Call SetFieldValue("TOTAL_PRICE_RT", NVLD(Rs("TOTAL_PRICE_RT"), 0))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      
      Call SetFieldValue("BRANCH_CODE", NVLS(Rs("BRANCH_CODE"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("APAR_CODE", NVLS(Rs("APAR_CODE"), ""))
      Call SetFieldValue("APAR_NAME", NVLS(Rs("APAR_NAME"), ""))
      Call SetFieldValue("EMPLOYEE_CODE", NVLS(Rs("EMPLOYEE_CODE"), ""))
      Call SetFieldValue("EMPLOYEE_NAME", NVLS(Rs("EMPLOYEE_LONG_NAME"), "") & " " & NVLS(Rs("EMPLOYEE_LAST_NAME"), ""))
      
      Call SetFieldValue("STOCK_NO", NVLS(Rs("STOCK_NO"), ""))
      Call SetFieldValue("STOCK_DESC", NVLS(Rs("STOCK_DESC"), ""))
      
   ElseIf Ind = 2 Then
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT_RT", NVLD(Rs("TOTAL_AMOUNT_RT"), 0))
      Call SetFieldValue("TOTAL_PRICE_RT", NVLD(Rs("TOTAL_PRICE_RT"), 0))
   ElseIf Ind = 3 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 4 Then
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT_RT", NVLD(Rs("TOTAL_AMOUNT_RT"), 0))
      Call SetFieldValue("TOTAL_PRICE_RT", NVLD(Rs("TOTAL_PRICE_RT"), 0))
      
      Call SetFieldValue("BRANCH_CODE", NVLS(Rs("BRANCH_CODE"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("EMPLOYEE_CODE", NVLS(Rs("EMPLOYEE_CODE"), ""))
      Call SetFieldValue("EMPLOYEE_NAME", NVLS(Rs("EMPLOYEE_LONG_NAME"), "") & " " & NVLS(Rs("EMPLOYEE_LAST_NAME"), ""))
      
      Call SetFieldValue("STOCK_NO", NVLS(Rs("STOCK_NO"), ""))
      Call SetFieldValue("STOCK_DESC", NVLS(Rs("STOCK_DESC"), ""))
      
      Call SetFieldValue("UNIT_AMOUNT", NVLD(Rs("UNIT_AMOUNT"), 0))
   ElseIf Ind = 5 Then
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT_RT", NVLD(Rs("TOTAL_AMOUNT_RT"), 0))
      Call SetFieldValue("TOTAL_PRICE_RT", NVLD(Rs("TOTAL_PRICE_RT"), 0))
      
      Call SetFieldValue("STOCK_NO", NVLS(Rs("STOCK_NO"), ""))
      Call SetFieldValue("STOCK_DESC", NVLS(Rs("STOCK_DESC"), ""))
      
      Call SetFieldValue("UNIT_AMOUNT", NVLD(Rs("UNIT_AMOUNT"), 0))
   ElseIf Ind = 6 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 7 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
   ElseIf Ind = 8 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 9 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 10 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 11 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 12 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 13 Then
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("GROUP_COM_ID", NVLI(Rs("GROUP_COM_ID"), -1))
      Call SetFieldValue("TOTAL_AMOUNT", NVLD(Rs("TOTAL_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("TOTAL_AMOUNT_RT", NVLD(Rs("TOTAL_AMOUNT_RT"), 0))
      Call SetFieldValue("TOTAL_PRICE_RT", NVLD(Rs("TOTAL_PRICE_RT"), 0))
   ElseIf Ind = 14 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
   ElseIf Ind = 15 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 16 Then
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_AMOUNT_RT = NVLD(Rs("TOTAL_AMOUNT_RT"), 0)
      TOTAL_PRICE_RT = NVLD(Rs("TOTAL_PRICE_RT"), 0)
      
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
   ElseIf Ind = 17 Then
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      SALE_JOINT_CODE = NVLS(Rs("SALE_JOINT_CODE"), "")
      SALE_LONG_NAME = NVLS(Rs("SALE_LONG_NAME"), "")
      SALE_LAST_NAME = NVLS(Rs("SALE_LAST_NAME"), "")
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      JOINT_CODE = NVLS(Rs("JOINT_CODE"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_AMOUNT = NVLS(Rs("UNIT_AMOUNT"), "")
   ElseIf Ind = 18 Then
      EMPLOYEE_CODE = NVLS(Rs("EMPLOYEE_CODE"), "")
      SALE_JOINT_CODE = NVLS(Rs("SALE_JOINT_CODE"), "")
      SALE_LONG_NAME = NVLS(Rs("SALE_LONG_NAME"), "")
      SALE_LAST_NAME = NVLS(Rs("SALE_LAST_NAME"), "")
   End If
End Sub
Public Sub RegisterFields()
   TableName = "TAGET_DETAIL"
   SequenceName = "TAGET_DETAIL_SEQ"
   PKField = "TAGET_DETAIL_ID"
   
   If MasterInd = 1 Then
      Call AddField("TAGET_DETAIL_ID", INT_TYPE, -1, ID_CAT)
      Call AddField("TAGET_ID", INT_TYPE, 0, DATA_CAT)
      Call AddField("BRANCH_ID", INT_TYPE, 0, DATA_CAT)
      Call AddField("APAR_MAS_ID", INT_TYPE, 0, DATA_CAT)
      Call AddField("EMP_ID", INT_TYPE, 0, DATA_CAT)
   
      Call AddField("TOTAL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TOTAL_AMOUNT_RT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("TOTAL_PRICE_RT", MONEY_TYPE, 0, DATA_CAT)
      Call AddField("PART_ITEM_ID", INT_TYPE, 0, DATA_CAT)
   
      Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
      Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
      Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
      Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

      Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
      Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
      
      Call AddField("BRANCH_CODE", STRING_TYPE, "", TEMP_CAT)
      Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("APAR_CODE", STRING_TYPE, "", TEMP_CAT)
      Call AddField("APAR_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("EMPLOYEE_CODE", STRING_TYPE, "", TEMP_CAT)
      Call AddField("EMPLOYEE_NAME", STRING_TYPE, "", TEMP_CAT)
      Call AddField("STOCK_NO", STRING_TYPE, "", TEMP_CAT)
      Call AddField("STOCK_DESC", STRING_TYPE, "", TEMP_CAT)
      
      Call AddField("YYYYMM", STRING_TYPE, "", TEMP_CAT)
      Call AddField("UNIT_AMOUNT", MONEY_TYPE, 0, TEMP_CAT)
      
      Call AddField("GROUP_COM_ID", INT_TYPE, -1, TEMP_CAT)
   End If
End Sub
Public Sub AddSearchAparTop(FromStr As String)     'สำหรับเอาเฉพาะลูกค้า TOP
   If InStr(1, FromStr, "APM") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (APM.APAR_MAS_ID = NM.APAR_MAS_ID) "
   End If
End Sub

