VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobItem"
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

Public JOB_ITEM_ID As Long
Public JOB_ID As Long

Public PART_ITEM_ID As Long
Public LOCATION_ID  As Long         ' คลัง

Public TX_TYPE As String
Public TX_AMOUNT As Double

Public LINK_ID As Long

Public UNIT_TRAN_ID As Long
Public UNIT_MULTIPLE As Double
'------------------------------------------------------------------------------------------------------------------------------------>

Public KEY_LOOKUP As String
Public KEY_ID As Long

Public ID As Long

Public PART_NO As String
Public PART_DESC As String
Public LOCATION_NO As String
Public LOCATION_NAME As String

Public UNIT_TRAN_NAME As String

Public UNIT_ID   As Long
Public UNIT_NAME As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_STOCK_NO As String
Public TO_STOCK_NO As String
Public FROM_LOCATION_NO As String
Public TO_LOCATION_NO As String

Public JOB_NO As String
Public JOB_DATE As Date
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
Dim TempStr  As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT JI.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE LOCATION_NO,MR1.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME UNIT_TRAN_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (JI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (JI.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (JI.UNIT_TRAN_ID = MR2.KEY_ID) "                          'หน่วยที่ผู้ใช้เปลี่ยนได้ ซึ่งถ้าไม่เปลี่ยนจะเป็น UNIT_ID หรือหน่วยใหญ่ ของสินค้านั้นๆ
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "                   'หน่วยย่อยที่ประจำสินค้านั้นๆ
   ElseIf Ind = 2 Then              'CReportProduct001
      SelectStr1 = "SELECT JI.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE LOCATION_NO,MR1.KEY_NAME LOCATION_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME UNIT_TRAN_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_NAME UNIT_NAME "
      SelectStr1 = SelectStr1 & ",J.JOB_NO,J.JOB_DATE "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_ITEM JI "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (JI.JOB_ID = J.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (JI.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (JI.LOCATION_ID = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (JI.UNIT_TRAN_ID = MR2.KEY_ID) "                          'หน่วยที่ผู้ใช้เปลี่ยนได้ ซึ่งถ้าไม่เปลี่ยนจะเป็น UNIT_ID หรือหน่วยใหญ่ ของสินค้านั้นๆ
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (SC.UNIT_CHANGE_ID = MR3.KEY_ID) "                   'หน่วยย่อยที่ประจำสินค้านั้นๆ
   End If
   
   If JOB_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JI.JOB_ITEM_ID = " & JOB_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JI.JOB_ITEM_ID = " & JOB_ITEM_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JI.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JI.JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (J.JOB_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (J.JOB_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY J.JOB_NO ASC,JI.TX_TYPE DESC,SC.STOCK_NO "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      JOB_ITEM_ID = NVLI(Rs("JOB_ITEM_ID"), -1)
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      UNIT_TRAN_ID = NVLI(Rs("UNIT_TRAN_ID"), -1)
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 2 Then
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      
      JOB_ITEM_ID = NVLI(Rs("JOB_ITEM_ID"), -1)
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      UNIT_TRAN_ID = NVLI(Rs("UNIT_TRAN_ID"), -1)
      UNIT_MULTIPLE = NVLD(Rs("UNIT_MULTIPLE"), 0)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      UNIT_TRAN_NAME = NVLS(Rs("UNIT_TRAN_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
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

   TempID1 = JOB_ID
   If JOB_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID4 = UNIT_TRAN_ID
   If UNIT_TRAN_ID <= 0 Then
      TempID4 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      
      SQL1 = "INSERT INTO JOB_ITEM " & _
                     "( " & _
                     "   JOB_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   TX_TYPE, " & _
                     "   TX_AMOUNT, " & _
                     "   UNIT_TRAN_ID, " & _
                     "   UNIT_MULTIPLE, " & _
                     "   LINK_ID, " & _
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
                     "'" & ChangeQuote(Trim(TX_TYPE)) & "'," & _
                     TX_AMOUNT & "," & _
                     TempID4 & "," & _
                     UNIT_MULTIPLE & "," & _
                     LINK_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ")  RETURNING JOB_ITEM_ID "
   Else 'Show edit
      SQL1 = "UPDATE JOB_ITEM SET " & _
                     "   PART_ITEM_ID = " & TempID2 & "," & _
                     "   LOCATION_ID = " & TempID3 & "," & _
                     "  TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "', " & _
                     "  TX_AMOUNT = " & TX_AMOUNT & ", " & _
                     "   UNIT_TRAN_ID = " & TempID4 & "," & _
                     "   UNIT_MULTIPLE = " & UNIT_MULTIPLE & "," & _
                     "   UNIT_MULTIPLE = " & LINK_ID & "," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_ITEM_ID = " & JOB_ITEM_ID
   End If
   ''debug.print (SQL1)
   If AddEditMode = SHOW_ADD Then
      JOB_ITEM_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM JOB_ITEM WHERE JOB_ITEM_ID = " & JOB_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

