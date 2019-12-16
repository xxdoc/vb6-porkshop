VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalanceVerifyDeTail"
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

Public BALANCE_VERIFY_DETAIL_ID As Long
Public BALANCE_VERIFY_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public PART_ITEM_AMOUNT As Double
'------------------------------------------------------------------------------------------------------------------------------------>

Public ID As Long

Public PART_NO As String
Public PART_DESC As String
Public LOCATION_NAME As String

Public BALANCE_VERIFY_DATE As Date
Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_STOCK_NO  As String
Public TO_STOCK_NO  As String
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
      SelectStr1 = "SELECT BVI.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO PART_NO,SC.STOCK_DESC PART_DESC "
      SelectStr1 = SelectStr1 & ",MR1.KEY_NAME LOCATION_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM BALANCE_VERIFY_DETAIL BVI "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = BVI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = BVI.LOCATION_ID) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SUM(BVI.PART_ITEM_AMOUNT) PART_ITEM_AMOUNT "
      SelectStr1 = SelectStr1 & ",BV.BALANCE_VERIFY_DATE,BVI.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM BALANCE_VERIFY_DETAIL BVI "
      FromStr = FromStr & "LEFT OUTER JOIN BALANCE_VERIFY BV ON (BV.BALANCE_VERIFY_ID = BVI.BALANCE_VERIFY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = BVI.PART_ITEM_ID) "
   End If
      
   If BALANCE_VERIFY_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BVI.BALANCE_VERIFY_DETAIL_ID = " & BALANCE_VERIFY_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BVI.BALANCE_VERIFY_DETAIL_ID = " & BALANCE_VERIFY_DETAIL_ID & ") "
      End If
   End If
   
   If BALANCE_VERIFY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BVI.BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BVI.BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BVI.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BVI.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BV.BALANCE_VERIFY_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BV.BALANCE_VERIFY_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(PART_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (SC.STOCK_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
      End If
   End If
   
   If Trim(FROM_STOCK_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO >= '" & ChangeQuote(Trim(FROM_STOCK_NO)) & "')"
      End If
   End If
   
   If Trim(TO_STOCK_NO) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (SC.STOCK_NO <= '" & ChangeQuote(Trim(TO_STOCK_NO)) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY BV.BALANCE_VERIFY_DATE,BVI.PART_ITEM_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   HaveOrder = True
   If Ind = 1 Then
      OrderByStr = "ORDER BY SC.STOCK_NO "
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
      BALANCE_VERIFY_DETAIL_ID = NVLI(Rs("BALANCE_VERIFY_DETAIL_ID"), -1)
      BALANCE_VERIFY_ID = NVLI(Rs("BALANCE_VERIFY_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_ITEM_AMOUNT = NVLD(Rs("PART_ITEM_AMOUNT"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      '------------------------------------------------------------------------------------------------------------------------------------>

      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
   ElseIf Ind = 2 Then
      PART_ITEM_AMOUNT = NVLD(Rs("PART_ITEM_AMOUNT"), 0)
      BALANCE_VERIFY_DATE = InternalDateToDate(NVLS(Rs("BALANCE_VERIFY_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
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

   TempID1 = BALANCE_VERIFY_ID
   If BALANCE_VERIFY_ID <= 0 Then
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
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BALANCE_VERIFY_DETAIL_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BALANCE_VERIFY_DETAIL " & _
                     "( " & _
                     "   BALANCE_VERIFY_DETAIL_ID, " & _
                     "   BALANCE_VERIFY_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   PART_ITEM_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     TempID3 & "," & _
                     PART_ITEM_AMOUNT & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         BALANCE_VERIFY_DETAIL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_VERIFY_DETAIL SET " & _
                     "   PART_ITEM_ID = " & TempID2 & "," & _
                     "   LOCATION_ID = " & TempID3 & "," & _
                     "  PART_ITEM_AMOUNT = " & PART_ITEM_AMOUNT & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_VERIFY_DETAIL_ID = " & BALANCE_VERIFY_DETAIL_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM BALANCE_VERIFY_DETAIL WHERE BALANCE_VERIFY_DETAIL_ID = " & BALANCE_VERIFY_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function
