VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployeeDealer"
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

Public EMPLOYEE_DEALER_ID As Long
Public EMP_ID As Long
Public YYYYMM As String
Public DEALER_TYPE As DEALER_TYPE_AREA
'------------------------------------------------------------------------------------------------------------------------------------>

Public ID As Long
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
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EMPDL.* "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM EMPLOYEE_DEALER EMPDL "
      
   End If
      
   If EMPLOYEE_DEALER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMPDL.EMPLOYEE_DEALER_ID = " & EMPLOYEE_DEALER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMPDL.EMPLOYEE_DEALER_ID = " & EMPLOYEE_DEALER_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMPDL.EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMPDL.EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   
   GroupByStr = ""
'   If Ind = 2 Then
'      GroupByStr = "GROUP BY BV.BALANCE_VERIFY_DATE,EMPDL.PART_ITEM_ID "
'   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      OrderByStr = "ORDER BY EMPDL.YYYYMM "
      OrderType = 2
   End If
   
   If Len(OrderByStr) > 0 Then
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
      EMPLOYEE_DEALER_ID = NVLI(Rs("EMPLOYEE_DEALER_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      DEALER_TYPE = NVLI(Rs("DEALER_TYPE"), -1)
      '------------------------------------------------------------------------------------------------------------------------------------>
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

   TempID1 = EMP_ID
   If EMP_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = DEALER_TYPE
   If DEALER_TYPE <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO EMPLOYEE_DEALER " & _
                     "( " & _
                     "   EMPLOYEE_DEALER_ID, " & _
                     "   EMP_ID, " & _
                     "   YYYYMM, " & _
                     "   DEALER_TYPE, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     TempID1 & "," & _
                     "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                     TempID2 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ")  RETURNING EMPLOYEE_DEALER_ID "
   Else 'Show edit
      SQL1 = "UPDATE EMPLOYEE_DEALER SET " & _
                     "   DEALER_TYPE = " & TempID2 & "," & _
                     "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMPLOYEE_DEALER_ID = " & EMPLOYEE_DEALER_ID
   End If
   '
   
   If AddEditMode = SHOW_ADD Then
      EMPLOYEE_DEALER_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EMPLOYEE_DEALER WHERE EMPLOYEE_DEALER_ID = " & EMPLOYEE_DEALER_ID
      m_Conn.Execute (SQL1)
End Function
