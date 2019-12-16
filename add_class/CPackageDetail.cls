VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPackageDetail"
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

Dim PACKAGE_DETAIL_ID As Long
Dim PACKAGE_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   PACKAGE_DETAIL_ID = Val(GetFieldValue("PACKAGE_DETAIL_ID"))
   PACKAGE_ID = Val(GetFieldValue("PACKAGE_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",PKG.PACKAGE_MASTER_FLAG "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (NM.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PACKAGE PKG ON (NM.PACKAGE_ID = PKG.PACKAGE_ID) "
   End If
   
   If PACKAGE_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PACKAGE_DETAIL_ID = " & PACKAGE_DETAIL_ID & ") " & vbCrLf
   End If
   
   If PACKAGE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PACKAGE_ID = " & PACKAGE_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY PACKAGE_DETAIL_ID "
      Else
         OrderByStr = "ORDER BY PACKAGE_DETAIL_ID "
      End If
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
      Call SetFieldValue("PACKAGE_DETAIL_ID", NVLI(Rs("PACKAGE_DETAIL_ID"), -1))
      Call SetFieldValue("PACKAGE_ID", NVLI(Rs("PACKAGE_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("PART_ITEM_COST", NVLD(Rs("PART_ITEM_COST"), 0))
      
      Call SetFieldValue("PART_NO", NVLS(Rs("PART_NO"), ""))
      Call SetFieldValue("PART_DESC", NVLS(Rs("PART_DESC"), ""))
      
      Call SetFieldValue("PACKAGE_MASTER_FLAG", NVLS(Rs("PACKAGE_MASTER_FLAG"), ""))
      
      Call SetFieldValue("PRO_FROM_DATE", InternalDateToDate(NVLS(Rs("PRO_FROM_DATE"), "")))
      Call SetFieldValue("PRO_TO_DATE", InternalDateToDate(NVLS(Rs("PRO_TO_DATE"), "")))
      Call SetFieldValue("PRO_ITEM_COST", NVLD(Rs("PRO_ITEM_COST"), 0))
      
   End If
End Sub
Public Sub RegisterFields()
   TableName = "PACKAGE_DETAIL"
   SequenceName = "PACKAGE_DETAIL_SEQ"
   PKField = "PACKAGE_DETAIL_ID"
   
   Call AddField("PACKAGE_DETAIL_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PACKAGE_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("PART_ITEM_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("PART_ITEM_COST", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("PRO_FROM_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("PRO_TO_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("PRO_ITEM_COST", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
   Call AddField("PART_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PART_DESC", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("PACKAGE_MASTER_FLAG", STRING_TYPE, "", TEMP_CAT)
   

   
End Sub
