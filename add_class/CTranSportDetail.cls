VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTranSportDetail"
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
Public TRANSPORT_DETAIL_ID As Long
Public DRIVER_ID As Long
Public CAR_LICENSE_ID As Long
Public TRANSPORTOR_ID As Long
Public TRANSPORT_DETAIL_DESC As String

Public TRANSPORT_PATH As String
Public CAR_TYPE As String
Public COST_PER_ROUND As Double
'----------------------------------------------------------------------------------------------------------------------------------------------
Public DRIVER_NAME As String
Public CAR_LICENSE_NAME As String
Public TRANSPORTOR_NAME As String

Public ORDER_BY  As Long
Public ORDER_TYPE  As Long
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
      
   SQL = "DELETE FROM TRANSPORT_DETAIL WHERE TRANSPORT_DETAIL_ID = " & TRANSPORT_DETAIL_ID
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
      SelectStr1 = "SELECT TSD.*"
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME DRIVER_NAME,MR2.KEY_NAME CAR_LICENSE_NAME,MR3.KEY_NAME TRANSPORTOR_NAME "
       
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM TRANSPORT_DETAIL TSD "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (TSD.DRIVER_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (TSD.CAR_LICENSE_ID = MR2.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (TSD.TRANSPORTOR_ID = MR3.KEY_ID) "
      
      WhereStr = ""
   End If
   
   If TRANSPORT_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (TSD.TRANSPORT_DETAIL_ID = " & TRANSPORT_DETAIL_ID & ") " & vbCrLf
   End If
   
   If TRANSPORTOR_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (TSD.TRANSPORTOR_ID = " & TRANSPORTOR_ID & ") " & vbCrLf
   End If
   
   If DRIVER_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (TSD.DRIVER_ID = " & DRIVER_ID & ") " & vbCrLf
   End If
   
   If CAR_LICENSE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (TSD.CAR_LICENSE_ID = " & CAR_LICENSE_ID & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY TSD.DRIVER_ID "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY TSD.CAR_LICENSE_ID "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY TSD.TRANSPORTOR_ID "
      ElseIf ORDER_BY = 4 Then
         OrderByStr = "ORDER BY MR3.KEY_NAME "
      Else
         OrderByStr = "ORDER BY TSD.TRANSPORT_DETAIL_ID "
      End If
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
      TRANSPORT_DETAIL_ID = NVLI(Rs("TRANSPORT_DETAIL_ID"), -1)
      DRIVER_ID = NVLI(Rs("DRIVER_ID"), -1)
      CAR_LICENSE_ID = NVLI(Rs("CAR_LICENSE_ID"), -1)
      TRANSPORTOR_ID = NVLI(Rs("TRANSPORTOR_ID"), -1)
      TRANSPORT_DETAIL_DESC = NVLS(Rs("TRANSPORT_DETAIL_DESC"), "")
      
      TRANSPORT_PATH = NVLS(Rs("TRANSPORT_PATH"), "")
      CAR_TYPE = NVLS(Rs("CAR_TYPE"), "")
      COST_PER_ROUND = NVLD(Rs("COST_PER_ROUND"), 0)
      
      DRIVER_NAME = NVLS(Rs("DRIVER_NAME"), "")
      CAR_LICENSE_NAME = NVLS(Rs("CAR_LICENSE_NAME"), "")
      TRANSPORTOR_NAME = NVLS(Rs("TRANSPORTOR_NAME"), "")
      
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = DRIVER_ID
   If DRIVER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = CAR_LICENSE_ID
   If CAR_LICENSE_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = TRANSPORTOR_ID
   If TRANSPORTOR_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      
      SQL1 = "INSERT INTO TRANSPORT_DETAIL " & _
                     "( " & _
                     "   DRIVER_ID, " & _
                     "   CAR_LICENSE_ID, " & _
                     "   TRANSPORTOR_ID, " & _
                     "   TRANSPORT_DETAIL_DESC, " & _
                     "   TRANSPORT_PATH, " & _
                     "   CAR_TYPE, " & _
                     "   COST_PER_ROUND, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                        TempID3 & "," & _
                        "'" & ChangeQuote(Trim(TRANSPORT_DETAIL_DESC)) & "'," & _
                        "'" & ChangeQuote(Trim(TRANSPORT_PATH)) & "'," & _
                        "'" & ChangeQuote(Trim(CAR_TYPE)) & "'," & _
                        COST_PER_ROUND & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         TRANSPORT_DETAIL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE TRANSPORT_DETAIL SET " & _
                     "    DRIVER_ID = " & TempID1 & ", " & _
                     "    CAR_LICENSE_ID = " & TempID2 & ", " & _
                     "    TRANSPORTOR_ID = " & TempID3 & ", " & _
                     "    TRANSPORT_DETAIL_DESC = '" & ChangeQuote(Trim(TRANSPORT_DETAIL_DESC)) & "', " & _
                     "    TRANSPORT_PATH = '" & ChangeQuote(Trim(TRANSPORT_PATH)) & "', " & _
                     "    CAR_TYPE = '" & ChangeQuote(Trim(CAR_TYPE)) & "', " & _
                     "    COST_PER_ROUND = " & COST_PER_ROUND & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE TRANSPORT_DETAIL_ID = " & TRANSPORT_DETAIL_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
