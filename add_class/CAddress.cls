VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAddress"
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

Dim ADDRESS_ID As Long
Dim COUNTRY_ID As Long
Dim APAR_MAS_ID As Long

Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
   
   ADDRESS_ID = Val(GetFieldValue("ADDRESS_ID"))
   COUNTRY_ID = Val(GetFieldValue("COUNTRY_ID"))
   
   APAR_MAS_ID = Val(GetFieldValue("APAR_MAS_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT AD.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " AD "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT ADR.*, CT.KEY_NAME COUNTRY_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ADDRESS ADR "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF CT ON (ADR.COUNTRY_ID = CT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_ADDRESS EA ON (ADR.ADDRESS_ID = EA.ADDRESS_ID) "
      
      WhereStr = " WHERE EA.ENTERPRISE_ADDRESS_ID IS NOT NULL "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT ADR.*, CT.KEY_NAME COUNTRY_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ADDRESS ADR "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF CT ON (ADR.COUNTRY_ID = CT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_ADDRESS CA ON (ADR.ADDRESS_ID = CA.ADDRESS_ID) "
      
      WhereStr = " WHERE CA.APAR_ADDRESS_ID IS NOT NULL "
   End If
   
   If ADDRESS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (AD.ADDRESS_ID = " & ADDRESS_ID & ") " & vbCrLf
   End If
   
   If APAR_MAS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APAR_MAS_ID = " & APAR_MAS_ID & ") "
      Else
         WhereStr = WhereStr & TempStr & "AND (APAR_MAS_ID = " & APAR_MAS_ID & ") "
      End If
   End If
   
   If COUNTRY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (AD.COUNTRY_ID = " & COUNTRY_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY HOME "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY ADDRESS_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY ADDRESS_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   'debug.print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("ADDRESS_ID", NVLI(Rs("ADDRESS_ID"), -1))
      Call SetFieldValue("COUNTRY_ID", NVLI(Rs("COUNTRY_ID"), -1))
      Call SetFieldValue("HOME", NVLS(Rs("HOME"), ""))
      Call SetFieldValue("SOI", NVLS(Rs("SOI"), ""))
      Call SetFieldValue("MOO", NVLS(Rs("MOO"), ""))
      Call SetFieldValue("VILLAGE", NVLS(Rs("VILLAGE"), ""))
      Call SetFieldValue("ROAD", NVLS(Rs("ROAD"), ""))
      Call SetFieldValue("DISTRICT", NVLS(Rs("DISTRICT"), ""))
      Call SetFieldValue("AMPHUR", NVLS(Rs("AMPHUR"), ""))
      Call SetFieldValue("PROVINCE", NVLS(Rs("PROVINCE"), ""))
      Call SetFieldValue("PHONE1", NVLS(Rs("PHONE1"), ""))
      Call SetFieldValue("PHONE2", NVLS(Rs("PHONE2"), ""))
      Call SetFieldValue("FAX1", NVLS(Rs("FAX1"), ""))
      Call SetFieldValue("FAX2", NVLS(Rs("FAX2"), ""))
      Call SetFieldValue("ZIPCODE", NVLS(Rs("ZIPCODE"), ""))
      Call SetFieldValue("BANGKOK_FLAG", NVLS(Rs("BANGKOK_FLAG"), ""))
      Call SetFieldValue("MAIN_FLAG", NVLS(Rs("MAIN_FLAG"), ""))
      Call SetFieldValue("SHOW_LOCATION_FLAG", NVLS(Rs("SHOW_LOCATION_FLAG"), "N"))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "ADDRESS"
   SequenceName = "ADDRESS_SEQ"
   PKField = "ADDRESS_ID"
   
   Call AddField("ADDRESS_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("COUNTRY_ID", INT_TYPE, "", DATA_CAT)
   Call AddField("HOME", STRING_TYPE, "", DATA_CAT)
   Call AddField("SOI", STRING_TYPE, "", DATA_CAT)
   Call AddField("MOO", STRING_TYPE, "", DATA_CAT)
   Call AddField("VILLAGE", STRING_TYPE, "", DATA_CAT)
   Call AddField("ROAD", STRING_TYPE, "", DATA_CAT)
   Call AddField("DISTRICT", STRING_TYPE, "", DATA_CAT)
   Call AddField("AMPHUR", STRING_TYPE, "", DATA_CAT)
   Call AddField("PROVINCE", STRING_TYPE, "", DATA_CAT)
   Call AddField("PHONE1", STRING_TYPE, "", DATA_CAT)
   Call AddField("PHONE2", STRING_TYPE, "", DATA_CAT)
   Call AddField("FAX1", STRING_TYPE, "", DATA_CAT)
   Call AddField("FAX2", STRING_TYPE, "", DATA_CAT)
   Call AddField("ZIPCODE", STRING_TYPE, "", DATA_CAT)
   Call AddField("BANGKOK_FLAG", STRING_TYPE, "", DATA_CAT)
   Call AddField("MAIN_FLAG", STRING_TYPE, "N", DATA_CAT)
   Call AddField("SHOW_LOCATION_FLAG", STRING_TYPE, "N", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("COUNTRY_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("ADDRESS_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("APAR_MAS_ID", INT_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, "", TEMP_CAT)
End Sub

Public Function PackAddress() As String
Dim AddressStr As String
Dim HOME As String
Dim SOI As String
Dim MOO As String
Dim VILLAGE As String
Dim ROAD As String
Dim DISTRICT As String
Dim AMPHUR As String
Dim PROVINCE As String
Dim ZIPCODE As String
Dim BANGKOK_FLAG As String
   
   HOME = GetFieldValue("HOME")
   SOI = GetFieldValue("SOI")
   MOO = GetFieldValue("MOO")
   VILLAGE = GetFieldValue("VILLAGE")
   ROAD = GetFieldValue("ROAD")
   DISTRICT = GetFieldValue("DISTRICT")
   AMPHUR = GetFieldValue("AMPHUR")
   PROVINCE = GetFieldValue("PROVINCE")
   ZIPCODE = GetFieldValue("ZIPCODE")
   BANGKOK_FLAG = GetFieldValue("BANGKOK_FLAG")
   
   AddressStr = ""
   
   If Trim(HOME) <> "" Then
      AddressStr = AddressStr & HOME & " "
   End If

   If Trim(SOI) <> "" Then
      AddressStr = AddressStr & "´." & SOI & " "
   End If

   If Trim(MOO) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË") & MOO & " "
   End If

   If Trim(VILLAGE) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË∫È“π") & VILLAGE & " "
   End If

   If Trim(ROAD) <> "" Then
      AddressStr = AddressStr & MapText("∂.") & ROAD & " "
   End If

   If Trim(DISTRICT) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("·¢«ß") & DISTRICT & " "
      Else
         AddressStr = AddressStr & MapText("µ”∫≈") & DISTRICT & " "
      End If
   End If

   If Trim(AMPHUR) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("‡¢µ") & AMPHUR & " "
      Else
         AddressStr = AddressStr & MapText("Õ”‡¿Õ") & AMPHUR & " "
      End If
   End If

   If Trim(PROVINCE) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & "" & PROVINCE & " "
      Else
         AddressStr = AddressStr & MapText("®.") & PROVINCE & " "
      End If
   End If

   If Trim(ZIPCODE) <> "" Then
      AddressStr = AddressStr & " " & ZIPCODE & " "
   End If

   PackAddress = AddressStr
End Function

