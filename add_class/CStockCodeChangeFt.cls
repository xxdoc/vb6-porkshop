VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockCodeChangeFt"
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
Public STOCK_CODE_CHANGE_FT_ID As Long
Public STOCK_CODE_ID As Long
Public BOX As Double                      'กล่อง
Public TRAY As Double                     'ถาด
Public PACK As Double               'แพ็ค
'----------------------------------------------------------------------------------------------------------------------------------------------
Public ORDER_BY  As Long
Public ORDER_TYPE  As Long
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
      
   SQL = "DELETE FROM STOCK_CODE_CHANGE_FT WHERE STOCK_CODE_CHANGE_FT_ID = " & STOCK_CODE_CHANGE_FT_ID
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
      SelectStr1 = "SELECT SCCF.*"
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE_CHANGE_FT SCCF "
            
      WhereStr = ""
   End If
   
   If STOCK_CODE_CHANGE_FT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SCCF.STOCK_CODE_CHANGE_FT_ID = " & STOCK_CODE_CHANGE_FT_ID & ") " & vbCrLf
   End If
   
   If STOCK_CODE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (SCCF.STOCK_CODE_ID = " & STOCK_CODE_ID & ") " & vbCrLf
   End If
      
   If Ind = 1 Then
      OrderByStr = "ORDER BY SCCF.STOCK_CODE_CHANGE_FT_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   ElseIf ORDER_TYPE = 2 Then
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      STOCK_CODE_CHANGE_FT_ID = NVLI(Rs("STOCK_CODE_CHANGE_FT_ID"), -1)
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      BOX = NVLD(Rs("BOX"), 0)
      TRAY = NVLD(Rs("TRAY"), 0)
      PACK = NVLD(Rs("PACK"), 0)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("STOCK_CODE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO STOCK_CODE_CHANGE_FT " & _
                     "( " & _
                     "   STOCK_CODE_ID, " & _
                     "   BOX, " & _
                     "   TRAY, " & _
                     "   PACK, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         STOCK_CODE_ID & "," & _
                         BOX & "," & _
                        TRAY & "," & _
                        PACK & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         'STOCK_CODE_CHANGE_FT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STOCK_CODE_CHANGE_FT SET " & _
                     "    BOX = " & BOX & ", " & _
                     "    TRAY = " & TRAY & ", " & _
                     "    PACK = " & PACK & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STOCK_CODE_CHANGE_FT_ID = " & STOCK_CODE_CHANGE_FT_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
