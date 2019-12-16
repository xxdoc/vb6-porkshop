VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRebateCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public Flag As String
' End of generic data member

Public REBATE_CONDITION_ID As Long
Public MASTER_VALID_ID As Long
Public CUSTOMER_TYPE As Long
Public STOCK_GROUP As Long

Public REBATE_LEVEL As Long
Public REBATE_PERCENT As Double

Public CUSTOMER_TYPE_NAME As String
Public STOCK_GROUP_NAME As String

Public OrderBy  As Long
Public OrderType As Long

Public VALID_FROM As Date
Public VALID_TO As Date

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr  As String


   If Ind = 1 Then
      SelectStr1 = "SELECT RBCDT.*  "
      SelectStr1 = SelectStr1 & ",CTMT.KEY_NAME CUSTOMER_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",STG.KEY_NAME STOCK_GROUP_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM REBATE_CONDITION RBCDT "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF CTMT ON (RBCDT.CUSTOMER_TYPE = CTMT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF STG ON (RBCDT.STOCK_GROUP = STG.KEY_ID) "
   ElseIf Ind = 2 Then                       'LoadDataRebateCondition
      SelectStr1 = "SELECT RBCDT.*  "
      
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM REBATE_CONDITION RBCDT "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_VALID MTVD ON (RBCDT.MASTER_VALID_ID = MTVD.MASTER_VALID_ID) "
   End If
   WhereStr = ""
   
   If REBATE_CONDITION_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (RBCDT.REBATE_CONDITION_ID = " & REBATE_CONDITION_ID & ")"
      Else
         WhereStr = WhereStr & "AND (RBCDT.REBATE_CONDITION_ID = " & REBATE_CONDITION_ID & ")"
      End If
   End If
   
   If MASTER_VALID_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (RBCDT.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      Else
         WhereStr = WhereStr & "AND (RBCDT.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      End If
   End If
   
   If VALID_FROM > 0 Then
      TempDate = DateToStringIntLow(Trim(VALID_FROM))
      If WhereStr = "" Then
         WhereStr = " WHERE (MTVD.VALID_FROM <=  '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (MTVD.VALID_FROM<= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If VALID_TO > 0 Then
      TempDate = DateToStringIntHi(Trim(VALID_TO))
      If WhereStr = "" Then
         WhereStr = " WHERE (MTVD.VALID_TO>= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (MTVD.VALID_TO>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RBCDT.REBATE_CONDITION_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY RBCDT.REBATE_CONDITION_ID "
      Else
         OrderByStr = "ORDER BY CTMT.KEY_CODE ASC,STG.KEY_CODE ASC,RBCDT.REBATE_LEVEL "
         OrderType = 1
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY RBCDT.CUSTOMER_TYPE ASC,RBCDT.STOCK_GROUP ASC,RBCDT.REBATE_LEVEL "
      OrderType = 1
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      REBATE_CONDITION_ID = NVLI(Rs("REBATE_CONDITION_ID"), -1)
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      
      CUSTOMER_TYPE = NVLI(Rs("CUSTOMER_TYPE"), -1)
      STOCK_GROUP = NVLI(Rs("STOCK_GROUP"), -1)
      
      REBATE_LEVEL = NVLI(Rs("REBATE_LEVEL"), -1)
      
      REBATE_PERCENT = NVLD(Rs("REBATE_PERCENT"), 0)
      
      CUSTOMER_TYPE_NAME = NVLS(Rs("CUSTOMER_TYPE_NAME"), "")
      STOCK_GROUP_NAME = NVLS(Rs("STOCK_GROUP_NAME"), "")
   ElseIf Ind = 2 Then
      REBATE_CONDITION_ID = NVLI(Rs("REBATE_CONDITION_ID"), -1)
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      
      CUSTOMER_TYPE = NVLI(Rs("CUSTOMER_TYPE"), -1)
      STOCK_GROUP = NVLI(Rs("STOCK_GROUP"), -1)
      
      REBATE_LEVEL = NVLI(Rs("REBATE_LEVEL"), -1)
      
      REBATE_PERCENT = NVLD(Rs("REBATE_PERCENT"), 0)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = CUSTOMER_TYPE
   If CUSTOMER_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = STOCK_GROUP
   If STOCK_GROUP <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO REBATE_CONDITION " & _
                     "( " & _
                     "   MASTER_VALID_ID, " & _
                     "   CUSTOMER_TYPE, " & _
                     "   STOCK_GROUP, " & _
                     "   REBATE_LEVEL, " & _
                     "   REBATE_PERCENT, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     MASTER_VALID_ID & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     REBATE_LEVEL & "," & _
                     REBATE_PERCENT & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_ID & "'" & _
                     ")  RETURNING REBATE_CONDITION_ID "
   Else
            SQL1 = "UPDATE REBATE_CONDITION SET " & _
                        "   CUSTOMER_TYPE = " & TempID1 & ", " & _
                        "   STOCK_GROUP = " & TempID2 & ", " & _
                        "   REBATE_LEVEL = " & REBATE_LEVEL & ", " & _
                        "   REBATE_PERCENT = " & REBATE_PERCENT & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = '" & glbUser.USER_ID & "' " & _
                        " WHERE REBATE_CONDITION_ID = " & REBATE_CONDITION_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      REBATE_CONDITION_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM REBATE_CONDITION WHERE REBATE_CONDITION_ID = " & REBATE_CONDITION_ID
      m_Conn.Execute (SQL1)
End Function


