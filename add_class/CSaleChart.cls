VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSaleChart"
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

Public SALE_CHART_ID As Long
Public MASTER_FROMTO_ID As Long
Public ORDER_ID As Long

Public PARENT_ID As Long
Public EMP_ID As Long
Public EMP_PERCENT  As Double
Public EMP_PARENT_PERCENT  As Double
'--------------------------------------------------------------------------------------------------------------------- >
Public MASTER_FROMTO_TYPE  As MASTER_COMMISSION_AREA

Public SALE_CODE As String
Public SALE_NAME  As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DEALER_TYPE As DEALER_TYPE_AREA
Public Level As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
Dim TempStr  As String

   If Ind = 1 Then
      SelectStr1 = "SELECT SCT.* "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_CHART SCT "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (SCT.EMP_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then                    'CReportCommission002
      SelectStr1 = "SELECT SCT.* "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_CHART SCT "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (SCT.EMP_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MF ON (SCT.MASTER_FROMTO_ID = MF.MASTER_FROMTO_ID) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(FROM_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(TO_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
   ElseIf Ind = 3 Then                    'CReportCommission003
      SelectStr1 = "SELECT SCT.* "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,EM.DEALER_TYPE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_CHART SCT "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (SCT.EMP_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MF ON (SCT.MASTER_FROMTO_ID = MF.MASTER_FROMTO_ID) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(FROM_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(TO_DATE)
      WhereStr = WhereStr & TempStr & " (MF.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (EM.DEALER_TYPE > 0) "
   End If
   
   If SALE_CHART_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SCT.SALE_CHART_ID = " & SALE_CHART_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SCT.SALE_CHART_ID = " & SALE_CHART_ID & ") "
      End If
   End If
   
   If MASTER_FROMTO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SCT.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SCT.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") "
      End If
   End If
   
   If MASTER_FROMTO_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MFT.MASTER_FROMTO_TYPE = " & MASTER_FROMTO_TYPE & ") " & vbCrLf
   End If
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY EM.EMP_CODE "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY SCT.MASTER_FROMTO_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY SCT.MASTER_FROMTO_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY EM.EMP_CODE "
   Else
      OrderByStr = "ORDER BY SCT.ORDER_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      SALE_CHART_ID = NVLI(Rs("SALE_CHART_ID"), -1)
      ORDER_ID = NVLI(Rs("ORDER_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_PERCENT = NVLD(Rs("EMP_PERCENT"), 0)
      EMP_PARENT_PERCENT = NVLD(Rs("EMP_PARENT_PERCENT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_LONG_NAME"), "") & " " & NVLS(Rs("SALE_LAST_NAME"), "")
   ElseIf Ind = 2 Then
      SALE_CHART_ID = NVLI(Rs("SALE_CHART_ID"), -1)
      ORDER_ID = NVLI(Rs("ORDER_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_PERCENT = NVLD(Rs("EMP_PERCENT"), 0)
      EMP_PARENT_PERCENT = NVLD(Rs("EMP_PARENT_PERCENT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_LONG_NAME"), "") & " " & NVLS(Rs("SALE_LAST_NAME"), "")
   ElseIf Ind = 3 Then
      SALE_CHART_ID = NVLI(Rs("SALE_CHART_ID"), -1)
      ORDER_ID = NVLI(Rs("ORDER_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_PERCENT = NVLD(Rs("EMP_PERCENT"), 0)
      EMP_PARENT_PERCENT = NVLD(Rs("EMP_PARENT_PERCENT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_LONG_NAME"), "") & " " & NVLS(Rs("SALE_LAST_NAME"), "")
      
      DEALER_TYPE = NVLI(Rs("DEALER_TYPE"), -1)
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
   
   TempID1 = MASTER_FROMTO_ID
   If MASTER_FROMTO_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = EMP_ID
   If EMP_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("SALE_CHART_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SALE_CHART " & _
                     "( " & _
                     "   MASTER_FROMTO_ID, " & _
                     "   ORDER_ID, " & _
                     "   PARENT_ID, " & _
                     "   EMP_ID, " & _
                     "   EMP_PERCENT, " & _
                     "   EMP_PARENT_PERCENT, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     TempID1 & "," & _
                     ORDER_ID & "," & _
                     TempID2 & "," & _
                     TempID3 & "," & _
                     EMP_PERCENT & "," & _
                     EMP_PARENT_PERCENT & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         'SALE_CHART_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SALE_CHART SET " & _
                     "   ORDER_ID = " & ORDER_ID & "," & _
                     "   PARENT_ID = " & TempID2 & "," & _
                     "   EMP_ID = " & TempID3 & "," & _
                     "  EMP_PERCENT = " & EMP_PERCENT & ", " & _
                     "  EMP_PARENT_PERCENT = " & EMP_PARENT_PERCENT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SALE_CHART_ID = " & SALE_CHART_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM SALE_CHART WHERE SALE_CHART_ID = " & SALE_CHART_ID
      m_Conn.Execute (SQL1)
     
End Function
