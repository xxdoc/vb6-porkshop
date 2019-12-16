VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKeyAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

'<---------------------------------------------------------------------------------------------------------------------

Public KEY_ACCOUNT_ID As Long
Public SALE_ID  As Long
Public KEY_ACCOUNT_DESC As String

'--------------------------------------------------------------------------------------------------------------------- >

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public SALE_CODE As String
Public SALE_LONG_NAME As String
Public SALE_LAST_NAME As String

Public KeyAccountDetail As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set KeyAccountDetail = New Collection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT BV.* "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM KEY_ACCOUNT BV "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (BV.SALE_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
   End If
   
   
   
   If KEY_ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BV.KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BV.KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID & ") "
      End If
   End If
   
   If SALE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BV.SALE_ID = " & SALE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BV.SALE_ID = " & SALE_ID & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BV.KEY_ACCOUNT_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY BV.KEY_ACCOUNT_DATE "
     Else
         OrderByStr = "ORDER BY BV.KEY_ACCOUNT_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      KEY_ACCOUNT_ID = NVLI(Rs("KEY_ACCOUNT_ID"), -1)
      KEY_ACCOUNT_DESC = NVLS(Rs("KEY_ACCOUNT_DESC"), "")
      SALE_ID = NVLI(Rs("SALE_ID"), -1)
      
      SALE_LONG_NAME = NVLS(Rs("SALE_LONG_NAME"), "")
      SALE_LAST_NAME = NVLS(Rs("SALE_LAST_NAME"), "")
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("KEY_ACCOUNT_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO KEY_ACCOUNT " & _
                     "( " & _
                     "   KEY_ACCOUNT_ID, " & _
                     "   KEY_ACCOUNT_DESC, " & _
                     "   SALE_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     "'" & ChangeQuote(Trim(KEY_ACCOUNT_DESC)) & "'," & _
                     SALE_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         KEY_ACCOUNT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE KEY_ACCOUNT SET " & _
                     "  KEY_ACCOUNT_DESC = '" & ChangeQuote(Trim(KEY_ACCOUNT_DESC)) & "', " & _
                     "   SALE_ID = " & SALE_ID & "," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID
   End If
   
   If AddEditMode = SHOW_ADD Or AddEditMode = SHOW_EDIT Then
       m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM KEY_ACCOUNT_DETAIL WHERE KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID
      m_Conn.Execute (SQL1)
     
      SQL1 = "DELETE FROM KEY_ACCOUNT WHERE KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID
      m_Conn.Execute (SQL1)
      
End Function
Private Sub Class_Terminate()
   Set KeyAccountDetail = Nothing
End Sub
