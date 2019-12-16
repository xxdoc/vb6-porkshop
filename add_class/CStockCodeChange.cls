VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockCodeChange"
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
Public STOCK_CODE_CHANGE_ID As Long
Public STOCK_CODE_ID As Long
Public UNIT_CHANGE_ID As Long
Public UNIT_CHANGE_AMOUNT As Double           'จำนวน หน่วยใหญ่ต่อหน่วยย่อย
Public CUSTOMER_ID As Long
'----------------------------------------------------------------------------------------------------------------------------------------------
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public UNIT_CHANGE_NAME As String

Public ORDER_BY  As Long
Public ORDER_TYPE  As Long
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
End Sub
Public Sub DeleteData()
Dim SQL As String
      
   SQL = "DELETE FROM STOCK_CODE_CHANGE WHERE STOCK_CODE_CHANGE_ID = " & STOCK_CODE_CHANGE_ID
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
      SelectStr1 = "SELECT NM.*, NM2.LONG_NAME CUSTOMER_NAME, APM.APAR_CODE CUSTOMER_CODE "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME UNIT_CHANGE_NAME "
       
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE_CHANGE NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (NM.UNIT_CHANGE_ID = MR2.KEY_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (NM.CUSTOMER_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT NM.STOCK_CODE_ID,NM.UNIT_CHANGE_ID,NM.UNIT_CHANGE_AMOUNT "

      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM STOCK_CODE_CHANGE NM "
      
      WhereStr = ""
   End If
   
   If STOCK_CODE_CHANGE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.STOCK_CODE_CHANGE_ID = " & STOCK_CODE_CHANGE_ID & ") " & vbCrLf
   End If
   
   If STOCK_CODE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.STOCK_CODE_ID = " & STOCK_CODE_ID & ") " & vbCrLf
   End If
      
   If Ind = 1 Then
      OrderByStr = "ORDER BY NM.STOCK_CODE_CHANGE_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY NM.STOCK_CODE_CHANGE_ID "
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
      STOCK_CODE_CHANGE_ID = NVLI(Rs("STOCK_CODE_CHANGE_ID"), -1)
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      UNIT_CHANGE_AMOUNT = NVLD(Rs("UNIT_CHANGE_AMOUNT"), 0)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      UNIT_CHANGE_NAME = NVLS(Rs("UNIT_CHANGE_NAME"), "")
   ElseIf Ind = 2 Then
      STOCK_CODE_ID = NVLI(Rs("STOCK_CODE_ID"), -1)
      UNIT_CHANGE_ID = NVLI(Rs("UNIT_CHANGE_ID"), -1)
      UNIT_CHANGE_AMOUNT = NVLD(Rs("UNIT_CHANGE_AMOUNT"), 0)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   TempID1 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("STOCK_CODE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO STOCK_CODE_CHANGE " & _
                     "( " & _
                     "   STOCK_CODE_ID, " & _
                     "   UNIT_CHANGE_ID, " & _
                     "   UNIT_CHANGE_AMOUNT, " & _
                     "   CUSTOMER_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         STOCK_CODE_ID & "," & _
                         UNIT_CHANGE_ID & "," & _
                        UNIT_CHANGE_AMOUNT & "," & _
                        TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         STOCK_CODE_CHANGE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STOCK_CODE_CHANGE SET " & _
                     "    UNIT_CHANGE_ID = " & UNIT_CHANGE_ID & ", " & _
                     "    UNIT_CHANGE_AMOUNT = " & UNIT_CHANGE_AMOUNT & ", " & _
                     "    CUSTOMER_ID = " & TempID1 & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STOCK_CODE_CHANGE_ID = " & STOCK_CODE_CHANGE_ID
   End If
   
   If (ShowMode = SHOW_ADD) Or (ShowMode = SHOW_EDIT) Then
      m_ADOConn.Execute (SQL1)
   End If
End Function
