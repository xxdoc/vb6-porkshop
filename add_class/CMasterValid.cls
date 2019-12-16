VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public Flag As String
' End of generic data member

Public MASTER_VALID_ID As Long
Public MASTER_VALID_NO As String
Public MASTER_VALID_DESC As String
Public MASTER_VALID_TYPE As MASTER_REBATE_AREA

Public VALID_FROM As Date
Public VALID_TO As Date

Public OrderBy  As Long
Public OrderType As Long

Public CollRebateCondition As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Set CollRebateCondition = New Collection
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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr  As String

   If Ind = 1 Then
      SelectStr1 = "SELECT MTVD.*  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM MASTER_VALID MTVD "
   End If
   WhereStr = ""
   
   If MASTER_VALID_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (MTVD.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      Else
         WhereStr = WhereStr & "AND (MTVD.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      End If
   End If
   
   If MASTER_VALID_TYPE > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (MTVD.MASTER_VALID_TYPE = " & MASTER_VALID_TYPE & ")"
      Else
         WhereStr = WhereStr & "AND (MTVD.MASTER_VALID_TYPE = " & MASTER_VALID_TYPE & ")"
      End If
   End If
   
   If VALID_FROM > 0 Then
      TempDate = DateToStringIntLow(Trim(VALID_FROM))
      If WhereStr = "" Then
         WhereStr = " WHERE (MTVD.VALID_FROM <=  '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (MTVD.VALID_FROM<= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If VALID_TO > 0 Then
      TempDate = DateToStringIntHi(Trim(VALID_TO))
      If WhereStr = "" Then
         WhereStr = " WHERE (MTVD.VALID_TO>= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (MTVD.VALID_TO>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MTVD.MASTER_VALID_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MTVD.MASTER_VALID_ID "
   Else
      OrderByStr = "ORDER BY MTVD.MASTER_VALID_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      MASTER_VALID_NO = NVLS(Rs("MASTER_VALID_NO"), "")
      MASTER_VALID_DESC = NVLS(Rs("MASTER_VALID_DESC"), "")
      MASTER_VALID_TYPE = NVLI(Rs("MASTER_VALID_TYPE"), -1)
      
      VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
      VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO MASTER_VALID " & _
                     "( " & _
                     "   MASTER_VALID_NO, " & _
                     "   MASTER_VALID_DESC, " & _
                     "   MASTER_VALID_TYPE, " & _
                     "   VALID_FROM, " & _
                     "   VALID_TO, " & _
                     "   CREATE_DATE, " & " CREATE_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     "'" & ChangeQuote(Trim(MASTER_VALID_NO)) & "'," & _
                     "'" & ChangeQuote(Trim(MASTER_VALID_DESC)) & "'," & _
                     MASTER_VALID_TYPE & "," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(VALID_FROM))) & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(VALID_TO))) & "'," & _
                     "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ")  RETURNING MASTER_VALID_ID "
   Else
            SQL1 = "UPDATE MASTER_VALID SET " & _
                        "   MASTER_VALID_NO = '" & ChangeQuote(Trim(MASTER_VALID_NO)) & "', " & _
                        "   MASTER_VALID_DESC = '" & ChangeQuote(Trim(MASTER_VALID_DESC)) & "', " & _
                        "    MASTER_VALID_TYPE = " & MASTER_VALID_TYPE & ", " & _
                        "   VALID_FROM = '" & ChangeQuote(Trim(DateToStringInt(VALID_FROM))) & "', " & _
                        "   VALID_TO = '" & ChangeQuote(Trim(DateToStringInt(VALID_TO))) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                        " WHERE MASTER_VALID_ID = " & MASTER_VALID_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      MASTER_VALID_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM REBATE_CONDITION WHERE MASTER_VALID_ID = " & MASTER_VALID_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "UPDATE NETWORK_CHART SET PARENT_ID = NULL WHERE MASTER_VALID_ID = " & MASTER_VALID_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM NETWORK_CHART WHERE MASTER_VALID_ID = " & MASTER_VALID_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM MASTER_VALID WHERE MASTER_VALID_ID = " & MASTER_VALID_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CollRebateCondition = Nothing
End Sub
