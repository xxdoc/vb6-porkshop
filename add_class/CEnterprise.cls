VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnterprise"
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

Public ENTERPRISE_ID As Long
Public ENTERPRISE_NAME As String
Public SHORT_NAME As String

Public TAX_ID As String
Public NOTE As String

Public EnpAddresses As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set EnpAddresses = New Collection
End Sub

'In this class we will not handle any error or any transaction
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

   SelectStr1 = "SELECT ETP.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ENTERPRISE ETP "
   
   WhereStr = ""
   
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ETP.ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ETP.ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ETP.ENTERPRISE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ETP.ENTERPRISE_ID "
   Else
      OrderByStr = "ORDER BY ETP.ENTERPRISE_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
   SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
   ENTERPRISE_NAME = NVLS(Rs("ENTERPRISE_NAME"), "")
   
   TAX_ID = NVLS(Rs("TAX_ID"), "")
   NOTE = NVLS(Rs("NOTE"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE ENTERPRISE SET " & _
                     "   ENTERPRISE_NAME = '" & ChangeQuote(Trim(ENTERPRISE_NAME)) & "', " & _
                     "   SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "', " & _
                     "   TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID
                    
   m_Conn.Execute (SQL1)
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM ENTERPRISE_ADDRESS WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM ENTERPRISE WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
      m_Conn.Execute (SQL1)
End Function
Private Sub Class_Terminate()
   Set EnpAddresses = Nothing
End Sub


