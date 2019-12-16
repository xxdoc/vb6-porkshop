VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNetworkChart"
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

Public NETWORK_CHART_ID As Long
Public MASTER_VALID_ID As Long
Public PARENT_ID As Long
Public CUSTOMER_ID As Long
Public ORDER_ID As Long
Public HOLD_FLAG As String

Public OrderBy  As Long
Public OrderType As Long

Public OLD_PK As Long

Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String

Public VALID_FROM As Date
Public VALID_TO As Date

Public Level As Long
Public SHOW_FLAG As String

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
      SelectStr1 = "SELECT NWC.*  "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE CUSTOMER_CODE,APM.APAR_NAME CUSTOMER_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM NETWORK_CHART NWC "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NWC.CUSTOMER_ID = APM.APAR_MAS_ID ) "
   ElseIf Ind = 2 Then                                      'CReportRebate001
      SelectStr1 = "SELECT NWC.*  "
      SelectStr1 = SelectStr1 & ",APM.APAR_CODE CUSTOMER_CODE,APM.APAR_NAME CUSTOMER_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM NETWORK_CHART NWC "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NWC.CUSTOMER_ID = APM.APAR_MAS_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_VALID MTVD ON (NWC.MASTER_VALID_ID = MTVD.MASTER_VALID_ID) "
   End If
   WhereStr = ""
   
   If NETWORK_CHART_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (NWC.NETWORK_CHART_ID = " & NETWORK_CHART_ID & ")"
      Else
         WhereStr = WhereStr & "AND (NWC.NETWORK_CHART_ID = " & NETWORK_CHART_ID & ")"
      End If
   End If
   
   If MASTER_VALID_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (NWC.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      Else
         WhereStr = WhereStr & "AND (NWC.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
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
         OrderByStr = "ORDER BY NWC.NETWORK_CHART_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY NWC.NETWORK_CHART_ID "
      Else
         OrderByStr = "ORDER BY NWC.PARENT_ID ASC,NWC.ORDER_ID "
         OrderType = 1
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY NWC.PARENT_ID ASC,NWC.ORDER_ID "
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
      NETWORK_CHART_ID = NVLI(Rs("NETWORK_CHART_ID"), -1)
      OLD_PK = NETWORK_CHART_ID
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      HOLD_FLAG = NVLS(Rs("HOLD_FLAG"), "")
      
      ORDER_ID = NVLI(Rs("ORDER_ID"), -1)
   ElseIf Ind = 2 Then
      NETWORK_CHART_ID = NVLI(Rs("NETWORK_CHART_ID"), -1)
      OLD_PK = NETWORK_CHART_ID
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      HOLD_FLAG = NVLS(Rs("HOLD_FLAG"), "")
      
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO NETWORK_CHART " & _
                     "( " & _
                     "   MASTER_VALID_ID, " & _
                     "   PARENT_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   ORDER_ID, " & _
                     "   HOLD_FLAG, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     MASTER_VALID_ID & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     ORDER_ID & "," & _
                     "'" & ChangeQuote(Trim(HOLD_FLAG)) & "'," & _
                     "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ")  RETURNING NETWORK_CHART_ID "
   Else
            SQL1 = "UPDATE NETWORK_CHART SET " & _
                        "   PARENT_ID = " & TempID1 & ", " & _
                        "   CUSTOMER_ID = " & TempID2 & ", " & _
                        "    ORDER_ID = " & ORDER_ID & ", " & _
                        "    HOLD_FLAG = '" & ChangeQuote(Trim(HOLD_FLAG)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                        " WHERE NETWORK_CHART_ID = " & NETWORK_CHART_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      NETWORK_CHART_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM NETWORK_CHART WHERE NETWORK_CHART_ID = " & NETWORK_CHART_ID
      m_Conn.Execute (SQL1)
End Function
