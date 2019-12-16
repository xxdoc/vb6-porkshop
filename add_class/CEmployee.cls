VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_ADOConn As ADODB.Connection

Public EMP_ID As Long
Public EMP_CODE As String
Public EMP_NAME  As String
Public MAINSALE_FLAG As String
'------------------------------------------------------------------------------------------------------------------------------------------------
Public FROM_SALE_CODE As String
Public TO_SALE_CODE As String
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public KEY_ID As Long
Public KEY_LOOKUP As String

Public APAR_CODE As String
Public APAR_NAME As String
Public APAR_SHORT_NAME As String
Public APAR_TYPE_NAME As String
Public APAR_GROUP_NAME As String
Public BRANCH_CODE As String
Public BRANCH_NAME As String

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
   SQL = "DELETE FROM EMPLOYEE WHERE EMP_ID = " & EMP_ID
   Call m_ADOConn.Execute(SQL)
End Sub
'Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, MustCount As Boolean)
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
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EMPLOYEE EM "
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EM.EMP_ID,EM.EMP_CODE, EM.EMP_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EMPLOYEE EM "
      
      
      
      WhereStr = ""
      
  ElseIf Ind = 3 Then
      SelectStr1 = "SELECT APM.APAR_CODE,APM.APAR_NAME "
      SelectStr1 = SelectStr1 & ",APT.KEY_NAME  APAR_TYPE_NAME  "
      SelectStr1 = SelectStr1 & ",APG.KEY_NAME  APAR_GROUP_NAME  "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE BRANCH_CODE ,MR.KEY_NAME BRANCH_NAME "    'ชื่อสาขา"
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EMPLOYEE EM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (MR.PARENT_EX_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MR.PARENT_EX_ID2 = APM.APAR_MAS_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APT ON (APM.APAR_TYPE = APT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF APG ON (APG.KEY_ID = APT.PARENT_ID) "
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.EMP_ID = " & EMP_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(EMP_CODE)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(EM.EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (EM.EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
      End If
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
   
   If Trim(EMP_NAME) <> "" Then
      If WildCard(EMP_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "EMP_NAME", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "EMP_NAME", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
         End If
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EM.EMP_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.EMP_NAME "
      Else
         OrderByStr = "ORDER BY EM.EMP_CODE "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EM.EMP_CODE "
   ElseIf Ind = 3 Then
     If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EM.EMP_CODE "
          If ORDER_TYPE = 1 Then
            OrderByStr = OrderByStr & " ASC "
         ElseIf ORDER_TYPE = 2 Then
            OrderByStr = OrderByStr & " DESC "
         Else
            OrderByStr = OrderByStr & " ASC "
          End If
          OrderByStr = OrderByStr & " ,APM.APAR_CODE ASC,MR.KEY_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.EMP_NAME "
          If ORDER_TYPE = 1 Then
            OrderByStr = OrderByStr & " ASC "
         ElseIf ORDER_TYPE = 2 Then
            OrderByStr = OrderByStr & " DESC "
         Else
            OrderByStr = OrderByStr & " ASC "
          End If
          OrderByStr = OrderByStr & " ,APM.APAR_CODE ASC,MR.KEY_CODE "
      Else
         OrderByStr = "ORDER BY EM.EMP_CODE "
          If ORDER_TYPE = 1 Then
            OrderByStr = OrderByStr & " ASC "
         ElseIf ORDER_TYPE = 2 Then
            OrderByStr = OrderByStr & " DESC "
         Else
            OrderByStr = OrderByStr & " ASC "
          End If
          OrderByStr = OrderByStr & " ,APM.APAR_CODE ASC,MR.KEY_CODE "
      End If
      ORDER_TYPE = 1
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      
      MAINSALE_FLAG = NVLS(Rs("MAINSALE_FLAG"), "")
      
      KEY_ID = NVLI(Rs("EMP_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("EMP_CODE"), "")
   ElseIf Ind = 2 Then
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
    ElseIf Ind = 3 Then
      APAR_CODE = NVLS(Rs("APAR_CODE"), "")
      APAR_NAME = NVLS(Rs("APAR_NAME"), "")
      
      APAR_TYPE_NAME = NVLS(Rs("APAR_TYPE_NAME"), "")
      APAR_GROUP_NAME = NVLS(Rs("APAR_GROUP_NAME"), "")
            
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      BRANCH_NAME = NVLS(Rs("BRANCH_NAME"), "")
      
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
   End If
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMPLOYEE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMPLOYEE " & _
                     "( " & _
                     "   EMP_ID, " & _
                     "   EMP_CODE, " & _
                     "   EMP_NAME, " & _
                     "   MAINSALE_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                        ID & "," & _
                        "'" & ChangeQuote(Trim(EMP_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(EMP_NAME)) & "'," & _
                        "'" & ChangeQuote(Trim(MAINSALE_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMPLOYEE SET " & _
                     "    EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "', " & _
                     "    EMP_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "', " & _
                     "    MAINSALE_FLAG = '" & ChangeQuote(Trim(MAINSALE_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_ID = " & EMP_ID
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
