VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTagetJobDetail"
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
   
Public TAGET_JOB_DETAIL_ID As Long
Public TAGET_JOB_ID As Long

Public BATCH_NO As Long
Public OUTPUT_TYPE_ID  As Long
Public OUTPUT_AMOUNT As Double
'--------------------------------------------------------------------------------------------------------------------- >

Public ID As Long

Public OUTPUT_CODE As String
Public OUTPUT_DESC As String

Public INPUT_ID   As Long
Public MONTH_ID As Long
Public YEAR_NO As String
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
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT TJD.* "
      SelectStr1 = SelectStr1 & ",PDT.KEY_CODE OUTPUT_CODE,PDT.KEY_NAME OUTPUT_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM TAGET_JOB_DETAIL TJD "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF PDT ON (PDT.KEY_ID = TJD.OUTPUT_TYPE_ID) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT TJD.OUTPUT_TYPE_ID,TJD.OUTPUT_AMOUNT "
      SelectStr1 = SelectStr1 & ",TJ.INPUT_ID,TJD.BATCH_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM TAGET_JOB_DETAIL TJD "
      FromStr = FromStr & "LEFT OUTER JOIN TAGET_JOB TJ ON (TJ.TAGET_JOB_ID = TJD.TAGET_JOB_ID) "
   End If
   
   If TAGET_JOB_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TJD.TAGET_JOB_DETAIL_ID = " & TAGET_JOB_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TJD.TAGET_JOB_DETAIL_ID = " & TAGET_JOB_DETAIL_ID & ") "
      End If
   End If
   
   If TAGET_JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TJD.TAGET_JOB_ID = " & TAGET_JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TJD.TAGET_JOB_ID = " & TAGET_JOB_ID & ") "
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
     OrderByStr = "ORDER BY TJD.TAGET_JOB_DETAIL_ID "
   ElseIf Ind = 2 Then
      OrderType = 3
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   ElseIf OrderType = 3 Then
      
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TAGET_JOB_DETAIL_ID = NVLI(Rs("TAGET_JOB_DETAIL_ID"), -1)
      TAGET_JOB_ID = NVLI(Rs("TAGET_JOB_ID"), -1)
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
      OUTPUT_TYPE_ID = NVLI(Rs("OUTPUT_TYPE_ID"), -1)
      OUTPUT_AMOUNT = NVLD(Rs("OUTPUT_AMOUNT"), 0)
      
      OUTPUT_CODE = NVLS(Rs("OUTPUT_CODE"), "")
      OUTPUT_DESC = NVLS(Rs("OUTPUT_DESC"), "")
   ElseIf Ind = 2 Then
      OUTPUT_TYPE_ID = NVLI(Rs("OUTPUT_TYPE_ID"), -1)
      OUTPUT_AMOUNT = NVLD(Rs("OUTPUT_AMOUNT"), 0)
      
      INPUT_ID = NVLI(Rs("INPUT_ID"), -1)
      BATCH_NO = NVLI(Rs("BATCH_NO"), -1)
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
      
      SQL1 = "INSERT INTO TAGET_JOB_DETAIL " & _
                     "( " & _
                     "   TAGET_JOB_ID, " & _
                     "   BATCH_NO, " & _
                     "   OUTPUT_TYPE_ID, " & _
                     "   OUTPUT_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     TAGET_JOB_ID & "," & _
                     BATCH_NO & "," & _
                     OUTPUT_TYPE_ID & "," & _
                     OUTPUT_AMOUNT & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         'TAGET_JOB_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE TAGET_JOB_DETAIL SET " & _
                     "   BATCH_NO = " & BATCH_NO & "," & _
                     "   OUTPUT_TYPE_ID = " & OUTPUT_TYPE_ID & "," & _
                     "   OUTPUT_AMOUNT = " & OUTPUT_AMOUNT & "," & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE TAGET_JOB_DETAIL_ID = " & TAGET_JOB_DETAIL_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM TAGET_JOB_DETAIL WHERE TAGET_JOB_DETAIL_ID = " & TAGET_JOB_DETAIL_ID
      m_Conn.Execute (SQL1)
      
End Function
