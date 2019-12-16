VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalanceVerify"
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

Public BALANCE_VERIFY_ID As Long
Public BALANCE_VERIFY_NO As String
Public BALANCE_VERIFY_DESC As String
Public BALANCE_VERIFY_DATE As Date

'--------------------------------------------------------------------------------------------------------------------- >

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date

Public PART_NO As String

Public BalanceVerifyDetail As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set BalanceVerifyDetail = New Collection
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

   If Ind = 1 Then
      SelectStr1 = "SELECT BV.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM BALANCE_VERIFY BV "
      
   End If
   
   
   
   If BALANCE_VERIFY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BV.BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BV.BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID & ") "
      End If
   End If
   
   If Trim(BALANCE_VERIFY_NO) <> "" Then
      If WildCard(BALANCE_VERIFY_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(BV.BALANCE_VERIFY_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(BV.BALANCE_VERIFY_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BV.BALANCE_VERIFY_NO = '" & ChangeQuote(Trim(BALANCE_VERIFY_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (BV.BALANCE_VERIFY_NO = '" & ChangeQuote(Trim(BALANCE_VERIFY_NO)) & "')"
         End If
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BV.BALANCE_VERIFY_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BV.BALANCE_VERIFY_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BALANCE_VERIFY_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BALANCE_VERIFY_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   
   
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BV.BALANCE_VERIFY_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY BV.BALANCE_VERIFY_DATE "
     Else
         OrderByStr = "ORDER BY BV.BALANCE_VERIFY_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BALANCE_VERIFY_ID = NVLI(Rs("BALANCE_VERIFY_ID"), -1)
      BALANCE_VERIFY_NO = NVLS(Rs("BALANCE_VERIFY_NO"), "")
      BALANCE_VERIFY_DESC = NVLS(Rs("BALANCE_VERIFY_DESC"), "")
      BALANCE_VERIFY_DATE = InternalDateToDate(NVLS(Rs("BALANCE_VERIFY_DATE"), ""))
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
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BALANCE_VERIFY_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO BALANCE_VERIFY " & _
                     "( " & _
                     "   BALANCE_VERIFY_ID, " & _
                     "   BALANCE_VERIFY_NO, " & _
                     "   BALANCE_VERIFY_DESC, " & _
                     "   BALANCE_VERIFY_DATE, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     "'" & ChangeQuote(Trim(BALANCE_VERIFY_NO)) & "'," & _
                     "'" & ChangeQuote(Trim(BALANCE_VERIFY_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(BALANCE_VERIFY_DATE))) & "'," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         BALANCE_VERIFY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_VERIFY SET " & _
                     "   BALANCE_VERIFY_NO = '" & ChangeQuote(Trim(BALANCE_VERIFY_NO)) & "', " & _
                     "  BALANCE_VERIFY_DESC = '" & ChangeQuote(Trim(BALANCE_VERIFY_DESC)) & "', " & _
                     "   BALANCE_VERIFY_DATE= '" & ChangeQuote(Trim(DateToStringInt(BALANCE_VERIFY_DATE))) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM BALANCE_VERIFY_DETAIL WHERE BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID
      m_Conn.Execute (SQL1)
     
      SQL1 = "DELETE FROM BALANCE_VERIFY WHERE BALANCE_VERIFY_ID = " & BALANCE_VERIFY_ID
      m_Conn.Execute (SQL1)
      
End Function
Private Sub Class_Terminate()
   Set BalanceVerifyDetail = Nothing
End Sub
