VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAPARMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public Flag As String
Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE

Public CstAddresses As Collection

Public APAR_MAS_ID As Long
Public APAR_CODE As String
Public APAR_NAME As String
Public APAR_TYPE As Long
Public TAX_ID  As String
Public BIRTH_DATE As Date
Public EMAIL  As String
Public CREDIT  As Double
Public WEBSITE As String
Public BUSINESS_DESC As String

Public APAR_IND As Long
Public PACKAGE_ID  As Long
Public CANCEL_OUT_DOCUMENT As String

'------------------------------------------------------------------------------------------------------------------------------------------------
Public KEY_ID As Long
Public KEY_LOOKUP As String
Public FROM_APAR_CODE As String
Public TO_APAR_CODE As String
Public APAR_GROUP As Long
Public APAR_TYPE_NAME As String
Public APAR_GROUP_NAME  As String

Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public BRANCH_CODE As String
Public BRANCH_NAME As String
Public EMP_CODE As String
Public EMP_NAME As String

Public APAR_ADDRESS As String
Public MAIN_FLAG  As String
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set CstAddresses = New Collection
End Sub

Private Sub Class_Terminate()
   Set CstAddresses = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM APAR_ADDRESS WHERE APAR_MAS_ID = " & APAR_MAS_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM APAR_MAS WHERE APAR_MAS_ID = " & APAR_MAS_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, CountFlag As Boolean)
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

   
   If Ind = 1 Then
      SelectStr1 = "SELECT APM.*, MR2.KEY_NAME APAR_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (APM.APAR_TYPE = MR2.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT APM.APAR_MAS_ID,APM.APAR_CODE, APM.APAR_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT APM.APAR_CODE, APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",APT.KEY_NAME  APAR_TYPE_NAME  "
      SelectStr1 = SelectStr1 & ",APG.KEY_NAME  APAR_GROUP_NAME  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APT ON (APM.APAR_TYPE = APT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APG ON (APG.KEY_ID = APT.PARENT_ID) "
   
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT APM.APAR_CODE "
      SelectStr1 = SelectStr1 & ", APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ", ADR.* "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM APAR_MAS APM "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_ADDRESS APARAD ON (APARAD.APAR_MAS_ID = APM.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ADDRESS ADR ON (APARAD.ADDRESS_ID = ADR.ADDRESS_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APT ON (APM.APAR_TYPE = APT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APG ON (APG.KEY_ID = APT.PARENT_ID) "
   End If
   
   If APAR_MAS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_MAS_ID = " & APAR_MAS_ID & ") " & vbCrLf
   End If
   
   If APAR_IND > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_IND = " & APAR_IND & ") " & vbCrLf
   End If
   
   If Len(Trim(APAR_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(APAR_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(APAR_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "')"
      End If
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
   
   If Trim(APAR_NAME) <> "" Then
      If WildCard(APAR_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "APAR_NAME", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "APAR_NAME", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (APAR_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (APAR_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "')"
         End If
      End If
   End If
   
   If APAR_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APM.APAR_TYPE = " & APAR_TYPE & ") " & vbCrLf
   End If
   
   If APAR_GROUP > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (APT.PARENT_ID = " & APAR_GROUP & ") " & vbCrLf
   End If
   
    If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APM.APAR_NAME "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE "
      End If
   ElseIf Ind = 2 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APM.APAR_NAME "
      Else
         OrderByStr = "ORDER BY APM.APAR_CODE "
      End If
   ElseIf Ind = 3 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_NAME "
      Else
         OrderByStr = "ORDER BY APG.KEY_CODE,APT.KEY_CODE,APM.APAR_CODE "
      End If
   ElseIf Ind = 5 Then
         OrderByStr = "ORDER BY APM.APAR_CODE "
         ORDER_TYPE = 1
   End If
   
      If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
'   Debug.Print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, CountFlag, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_TYPE = NVLI(Rs("APAR_TYPE"), -1)
      CREDIT = NVLD(Rs("CREDIT"), 0)
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
      BUSINESS_DESC = NVLS(Rs("BUSINESS_DESC"), "")
      EMAIL = NVLS(Rs("EMAIL"), "")
      WEBSITE = NVLS(Rs("WEBSITE"), "")
      APAR_IND = NVLD(Rs("APAR_IND"), -1)
      PACKAGE_ID = NVLD(Rs("PACKAGE_ID"), -1)
      
      
      KEY_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      
      CANCEL_OUT_DOCUMENT = NVLS(Rs("CANCEL_OUT_DOCUMENT"), "")
   ElseIf Ind = 2 Then
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
   ElseIf Ind = 3 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
   ElseIf Ind = 5 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      MAIN_FLAG = NVLS(Rs("MAIN_FLAG"), "")
      
      Call PackAddress(Rs)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID2 = APAR_TYPE
   If APAR_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID4 = PACKAGE_ID
   If PACKAGE_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("APAR_MAS_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO APAR_MAS " & _
                     "( " & _
                     "   APAR_MAS_ID, " & _
                     "   APAR_CODE, " & _
                     "   APAR_NAME, " & _
                     "   APAR_TYPE, " & _
                     "   TAX_ID, " & _
                     "   BIRTH_DATE, " & _
                     "   EMAIL, " & _
                     "   CREDIT, " & _
                     "   WEBSITE, " & _
                     "   BUSINESS_DESC, " & _
                     "   APAR_IND, " & _
                     "   PACKAGE_ID, " & _
                     "   CANCEL_OUT_DOCUMENT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "    '
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                        ID & "," & _
                        "'" & ChangeQuote(Trim(APAR_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(APAR_NAME)) & "'," & _
                        TempID2 & "," & _
                        "'" & ChangeQuote(Trim(TAX_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                        CREDIT & "," & _
                        "'" & ChangeQuote(Trim(WEBSITE)) & "'," & _
                        "'" & ChangeQuote(Trim(BUSINESS_DESC)) & "'," & _
                        APAR_IND & "," & _
                        TempID4 & "," & _
                        "'" & ChangeQuote(Trim(CANCEL_OUT_DOCUMENT)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ") "    '
         APAR_MAS_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE APAR_MAS SET " & _
                     "    APAR_CODE = '" & ChangeQuote(Trim(APAR_CODE)) & "', " & _
                     "    APAR_NAME = '" & ChangeQuote(Trim(APAR_NAME)) & "', " & _
                     "    APAR_TYPE = " & TempID2 & ", " & _
                     "    CREDIT = " & CREDIT & ", " & _
                     "    TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   BIRTH_DATE= '" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "', " & _
                     "    EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                     "    WEBSITE = '" & ChangeQuote(Trim(WEBSITE)) & "', " & _
                     "    BUSINESS_DESC = '" & ChangeQuote(Trim(BUSINESS_DESC)) & "', " & _
                     "    APAR_IND = " & APAR_IND & ", " & _
                     "    PACKAGE_ID = " & TempID4 & ", " & _
                      "    CANCEL_OUT_DOCUMENT = '" & ChangeQuote(Trim(CANCEL_OUT_DOCUMENT)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE APAR_MAS_ID = " & APAR_MAS_ID
   End If

   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function

Private Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim I As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For I = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next I
    WhereStr = WhereStr & ")"
    
    GenerateSearchLike = WhereStr
End Function
Public Function PackAddress(Rs As ADODB.Recordset)
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
      
   HOME = NVLS(Rs("HOME"), "")
   SOI = NVLS(Rs("SOI"), "")
   MOO = NVLS(Rs("MOO"), "")
   VILLAGE = NVLS(Rs("VILLAGE"), "")
   ROAD = NVLS(Rs("ROAD"), "")
   DISTRICT = NVLS(Rs("DISTRICT"), "")
   AMPHUR = NVLS(Rs("AMPHUR"), "")
   PROVINCE = NVLS(Rs("PROVINCE"), "")
   ZIPCODE = NVLS(Rs("ZIPCODE"), "")
   BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "")
   
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
   APAR_ADDRESS = PackAddress
End Function

