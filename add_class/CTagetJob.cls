VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTagetJob"
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
Public TAGET_JOB_ID As Long
Public MONTH_ID As Long
Public YEAR_NO As String
Public INPUT_ID  As Long
Public INPUT_AMOUNT As Double
Public TAGET_JOB_DESC As String
'--------------------------------------------------------------------------------------------------------------------- >
Public ID As Long

Public INPUT_CODE As String
Public INPUT_DESC As String

Public TagetJobDetail As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set TagetJobDetail = New Collection
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
      SelectStr1 = "SELECT TJ.* "
      SelectStr1 = SelectStr1 & ",SC.STOCK_NO INPUT_CODE,SC.STOCK_DESC INPUT_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM TAGET_JOB TJ "
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (SC.STOCK_CODE_ID = TJ.INPUT_ID) "
   ElseIf Ind = 2 Then        'LoadTagetJobInputByType
      SelectStr1 = "SELECT TJ.INPUT_ID,TJ.INPUT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM TAGET_JOB TJ "
   End If
   
   If TAGET_JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TJ.TAGET_JOB_ID = " & TAGET_JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TJ.TAGET_JOB_ID = " & TAGET_JOB_ID & ") "
      End If
   End If
   
   If INPUT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TJ.INPUT_ID = " & INPUT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TJ.INPUT_ID = " & INPUT_ID & ") "
      End If
   End If
   
   If MONTH_ID > 0 Then
      If Len(WhereStr) > 0 Then
         WhereStr = WhereStr & "AND (TJ.MONTH_ID = " & MONTH_ID & ") "
      Else
         WhereStr = "WHERE (TJ.MONTH_ID = " & MONTH_ID & ") "
      End If
   End If
   
   If Trim(YEAR_NO) <> "" Then
      If WildCard(YEAR_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(TJ.YEAR_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(TJ.YEAR_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (TJ.YEAR_NO = '" & ChangeQuote(Trim(YEAR_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (TJ.YEAR_NO = '" & ChangeQuote(Trim(YEAR_NO)) & "')"
         End If
      End If
   End If
   
   GroupByStr = ""
'   If Ind = 2 Then
'   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
'   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY TJ.YEAR_NO ASC,TJ.MONTH_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SC.STOCK_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY SC.STOCK_DESC "
     Else
         OrderByStr = "ORDER BY TJ.TAGET_JOB_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TAGET_JOB_ID = NVLI(Rs("TAGET_JOB_ID"), -1)
      MONTH_ID = NVLI(Rs("MONTH_ID"), -1)
      YEAR_NO = NVLS(Rs("YEAR_NO"), "")
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      INPUT_AMOUNT = NVLD(Rs("INPUT_AMOUNT"), 0)
      
      TAGET_JOB_DESC = NVLS(Rs("TAGET_JOB_DESC"), "")
      
      INPUT_CODE = NVLS(Rs("INPUT_CODE"), "")
      INPUT_DESC = NVLS(Rs("INPUT_DESC"), "")
   ElseIf Ind = 2 Then
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      INPUT_AMOUNT = NVLD(Rs("INPUT_AMOUNT"), 0)
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
      Call glbDatabaseMngr.GetSeqID("TAGET_JOB_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO TAGET_JOB " & _
                     "( " & _
                     "   TAGET_JOB_ID, " & _
                     "   MONTH_ID, " & _
                     "   YEAR_NO, " & _
                     "   INPUT_ID, " & _
                     "   INPUT_AMOUNT, " & _
                     "   TAGET_JOB_DESC, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     MONTH_ID & "," & _
                     "'" & ChangeQuote(Trim(YEAR_NO)) & "'," & _
                     INPUT_ID & "," & _
                     INPUT_AMOUNT & "," & _
                     "'" & ChangeQuote(Trim(TAGET_JOB_DESC)) & "'," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         TAGET_JOB_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE TAGET_JOB SET " & _
                     "   MONTH_ID = " & MONTH_ID & "," & _
                     "   YEAR_NO = '" & ChangeQuote(Trim(YEAR_NO)) & "', " & _
                     "   INPUT_ID = " & INPUT_ID & "," & _
                     "   INPUT_AMOUNT = " & INPUT_AMOUNT & "," & _
                     "  TAGET_JOB_DESC = '" & ChangeQuote(Trim(TAGET_JOB_DESC)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE TAGET_JOB_ID = " & TAGET_JOB_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM TAGET_JOB_DETAIL WHERE TAGET_JOB_ID = " & TAGET_JOB_ID
      m_Conn.Execute (SQL1)
     
      SQL1 = "DELETE FROM TAGET_JOB WHERE TAGET_JOB_ID = " & TAGET_JOB_ID
      m_Conn.Execute (SQL1)
      
End Function
Private Sub Class_Terminate()
   Set TagetJobDetail = Nothing
End Sub
