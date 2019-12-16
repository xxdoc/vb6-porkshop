VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCapitalMovement"
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
Public CAPITAL_MOVEMENT_ID As Long
Public DOCUMENT_DATE As Date
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public CAPITAL_AMOUNT As Double
'--------------------------------------------------------------------------------------------------------------------- >

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date
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
   
   If Ind = 1 Then      'Normal and For LoadCapitalMovement
      SelectStr1 = "SELECT CM.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
   
   End If
   
   If CAPITAL_MOVEMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.STOCK_NO, PI.STOCK_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.CAPITAL_MOVEMENT_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CM.DOCUMENT_DATE "
     Else
         OrderByStr = "ORDER BY CM.CAPITAL_MOVEMENT_ID "
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
      CAPITAL_MOVEMENT_ID = NVLI(Rs("CAPITAL_MOVEMENT_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
'Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
   
   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("CAPITAL_MOVEMENT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO CAPITAL_MOVEMENT " & _
                     "( " & _
                     "   DOCUMENT_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   CAPITAL_AMOUNT, " & _
                     "   LOCATION_ID, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                     "" & TempID1 & "," & _
                     "" & CAPITAL_AMOUNT & "," & _
                     "" & TempID2 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         'CAPITAL_MOVEMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CAPITAL_MOVEMENT SET " & _
                     "   DOCUMENT_DATE= '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   PART_ITEM_ID = " & TempID1 & ", " & _
                     "   CAPITAL_AMOUNT = " & CAPITAL_AMOUNT & ", " & _
                     "   LOCATION_ID = " & TempID2 & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM CAPITAL_MOVEMENT WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
      m_Conn.Execute (SQL1)
      
End Function
Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
Dim TempStr  As String
   
   SQL1 = "DELETE FROM CAPITAL_MOVEMENT "
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PART_ITEM_ID = " & PART_ITEM_ID & ") " & vbCrLf
   End If
   
   SQL1 = SQL1 & WhereStr
   Call m_Conn.Execute(SQL1)
End Function
