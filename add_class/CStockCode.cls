VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

'----------------------------------------------------------------------------------------------------------------------------------------------
Public STOCK_CODE_ID As Long
Public STOCK_NO As String
Public STOCK_DESC As String
Public STOCK_TYPE As Long

Public UNIT_ID As Long                             ' ID หน่วยใหญ่
Public UNIT_AMOUNT As Double           'จำนวน หน่วยใหญ่ต่อหน่วยย่อย
Public UNIT_CHANGE_ID As Long                ' ID หน่วยย่อย
Public EXCEPTION_FLAG As String                 'ยก เว้นไม่ออกบิล
Public EXPENSE_FLAG As String                 'ค่าใช้จ่าย
'----------------------------------------------------------------------------------------------------------------------------------------------

Public GROUP_ID  As Long
Public UNIT_NAME  As String
Public UNIT_CHANGE_NAME  As String

Public STOCK_TYPE_CODE  As String
Public STOCK_TYPE_NAME  As String

Public STOCK_GROUP_CODE  As String
Public STOCK_GROUP_NAME  As String
'----------------------------------------------------------------------------------------------------------------------------------------------
Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String
Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public KEY_ID  As Long
Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
      
   SQL = "DELETE FROM STOCK_CODE WHERE STOCK_CODE_ID = " & STOCK_CODE_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
      SelectStr1 = "SELECT SC.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME, MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE SC "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (SC.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SC.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE SC "
      
      WhereStr = ""
   ElseIf Ind = 3 Then                             'CReportInventoryDoc10                 'CReportInventoryDoc1_1
      SelectStr1 = "SELECT SC.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME, MR2.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr1 = SelectStr1 & ", MR4.KEY_CODE STOCK_GROUP_CODE,MR4.KEY_NAME STOCK_GROUP_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE SC "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (SC.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (SC.UNIT_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (MR1.PARENT_ID = MR4.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 4 Then              'CReportInventoryDoc3_1 ,FOR LOAD TO CReportInventoryDoc3
      SelectStr1 = "SELECT SC.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_CHANGE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE SC "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (SC.STOCK_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "
      WhereStr = ""
   ElseIf Ind = 7 Then              'CReportInventoryDoc3_4
      SelectStr1 = "SELECT SC.*, MR1.KEY_CODE STOCK_TYPE_CODE,MR1.KEY_NAME STOCK_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE SC "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (SC.STOCK_TYPE = MR1.KEY_ID) "
      WhereStr = ""
   End If
   
   If STOCK_CODE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_CODE_ID = " & STOCK_CODE_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(STOCK_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "')"
      End If
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
   
   If Trim(STOCK_DESC) <> "" Then
      If WildCard(STOCK_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "STOCK_DESC", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "STOCK_DESC", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (STOCK_DESC = '" & ChangeQuote(Trim(STOCK_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (STOCK_DESC = '" & ChangeQuote(Trim(STOCK_DESC)) & "')"
         End If
      End If
   End If
      
   If STOCK_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.STOCK_TYPE = " & STOCK_TYPE & ") " & vbCrLf
   End If
   
   If UNIT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.UNIT_ID = " & UNIT_ID & ") " & vbCrLf
   End If
   
   If GROUP_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MR1.PARENT_ID = " & GROUP_ID & ") " & vbCrLf
   End If
      
   If Len(Trim(EXCEPTION_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "')"
   End If
   
   If Len(Trim(EXPENSE_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SC.EXPENSE_FLAG = '" & ChangeQuote(Trim(EXPENSE_FLAG)) & "')"
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE,SC.STOCK_NO  "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY MR1.KEY_CODE,SC.STOCK_NO "
      Else
         OrderByStr = "ORDER BY SC.STOCK_NO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY MR4.KEY_CODE ASC,MR1.KEY_CODE ASC,SC.STOCK_NO "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY MR1.KEY_CODE,SC.STOCK_NO  "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY SC.STOCK_NO  "
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
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      EXPENSE_FLAG = NVLS(Rs("EXPENSE_FLAG"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      KEY_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("STOCK_NO"), "")
   ElseIf Ind = 2 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
   ElseIf Ind = 3 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      UNIT_ID = NVLI(Rs("UNIT_ID"), -1)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      STOCK_TYPE = NVLI(Rs("STOCK_TYPE"), -1)
      
      UNIT_AMOUNT = NVLD(Rs("UNIT_AMOUNT"), 0)
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
      
      STOCK_GROUP_CODE = NVLS(Rs("STOCK_GROUP_CODE"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 4 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
      
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 7 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      STOCK_NO = NVLS(Rs("STOCK_NO"), "")
      STOCK_DESC = NVLS(Rs("STOCK_DESC"), "")
      
      STOCK_TYPE_CODE = NVLS(Rs("STOCK_TYPE_CODE"), "")
      STOCK_TYPE_NAME = NVLS(Rs("STOCK_TYPE_NAME"), "")
   End If
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4  As String
Dim TempID5  As String
Dim TempID6 As String
   
   TempID1 = UNIT_ID
   If UNIT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = STOCK_TYPE
   If STOCK_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = UNIT_CHANGE_ID
   If UNIT_CHANGE_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("STOCK_CODE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO STOCK_CODE " & _
                     "( " & _
                     "   STOCK_CODE_ID, " & _
                     "   STOCK_NO, " & _
                     "   STOCK_DESC, " & _
                     "   UNIT_ID, " & _
                     "   STOCK_TYPE, " & _
                     "   UNIT_AMOUNT, " & _
                     "   UNIT_CHANGE_ID, " & _
                     "   EXCEPTION_FLAG, "
         SQL1 = SQL1 & _
                     "   EXPENSE_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(STOCK_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(STOCK_DESC)) & "'," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        UNIT_AMOUNT & "," & _
                        TempID3 & ","
      SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(EXPENSE_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         STOCK_CODE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STOCK_CODE SET " & _
                     "    STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "', " & _
                     "    STOCK_DESC = '" & ChangeQuote(Trim(STOCK_DESC)) & "', " & _
                     "    UNIT_ID = " & TempID1 & ", " & _
                     "    STOCK_TYPE = " & TempID2 & ", " & _
                     "    UNIT_AMOUNT = " & UNIT_AMOUNT & ", " & _
                     "    UNIT_CHANGE_ID = " & TempID3 & ", " & _
                     "    EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "    EXPENSE_FLAG = '" & ChangeQuote(Trim(EXPENSE_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STOCK_CODE_ID = " & STOCK_CODE_ID
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
