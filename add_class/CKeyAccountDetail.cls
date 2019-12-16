VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKeyAccountDetail"
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

'<------------------------------------------------------------------------------------------------------------------------------------

Public KEY_ACCOUNT_DETAIL_ID As Long
Public KEY_ACCOUNT_ID As Long
Public CUSTOMER_ID As Long
Public BRANCH_ID As Long

'------------------------------------------------------------------------------------------------------------------------------------>
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public BRANCH_CODE As String
Public BRANCH_NAME As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
Dim TempStr As String
Dim HaveOrder As Boolean
   
   If Ind = 1 Then
      SelectStr1 = "SELECT KAI.* "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE CUSTOMER_CODE,NM2.SHORT_NAME CUSTOMER_NAME "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE BRANCH_CODE,MR1.KEY_NAME BRANCH_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM KEY_ACCOUNT_DETAIL KAI "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (APM.APAR_MAS_ID = KAI.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = KAI.BRANCH_ID) "
      
      WhereStr = ""
   End If
      
   If KEY_ACCOUNT_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (KAI.KEY_ACCOUNT_DETAIL_ID = " & KEY_ACCOUNT_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (KAI.KEY_ACCOUNT_DETAIL_ID = " & KEY_ACCOUNT_DETAIL_ID & ") "
      End If
   End If
   
   If KEY_ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (KAI.KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (KAI.KEY_ACCOUNT_ID = " & KEY_ACCOUNT_ID & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      'GroupByStr = "GROUP BY BV.KEY_ACCOUNT_DATE,KAI.BRANCH_ID,KAI.CUSTOMER_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   HaveOrder = True
   If Ind = 1 Then
      OrderByStr = "ORDER BY KAI.KEY_ACCOUNT_DETAIL_ID "
      OrderType = 1
   ElseIf Ind = 2 Then
      HaveOrder = False
   End If
   
   If HaveOrder Then
      If OrderType = 1 Then
         OrderByStr = OrderByStr & " ASC "
      ElseIf OrderType = 2 Then
         OrderByStr = OrderByStr & " DESC "
      Else
         OrderByStr = OrderByStr & " DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      '<------------------------------------------------------------------------------------------------------------------------------------
      KEY_ACCOUNT_DETAIL_ID = NVLI(Rs("KEY_ACCOUNT_DETAIL_ID"), -1)
      KEY_ACCOUNT_ID = NVLI(Rs("KEY_ACCOUNT_ID"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      '------------------------------------------------------------------------------------------------------------------------------------>
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      BRANCH_NAME = NVLS(Rs("BRANCH_NAME"), "")
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
Dim TempID5 As String
Dim TempID6 As String

   TempID1 = KEY_ACCOUNT_ID
   If KEY_ACCOUNT_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = BRANCH_ID
   If BRANCH_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      
      SQL1 = "INSERT INTO KEY_ACCOUNT_DETAIL " & _
                     "( " & _
                     "   KEY_ACCOUNT_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   BRANCH_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     TempID3 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         'KEY_ACCOUNT_DETAIL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE KEY_ACCOUNT_DETAIL SET " & _
                     "   CUSTOMER_ID = " & TempID2 & "," & _
                     "   BRANCH_ID = " & TempID3 & "," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE KEY_ACCOUNT_DETAIL_ID = " & KEY_ACCOUNT_DETAIL_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM KEY_ACCOUNT_DETAIL WHERE KEY_ACCOUNT_DETAIL_ID = " & KEY_ACCOUNT_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function

