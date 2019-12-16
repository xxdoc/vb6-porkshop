VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalanceVerifyDate"
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
Public BALANCE_VERIFY_DATE_ID As Long
Public BALANCE_VERIFY_DETAIL_ID As Long
Public IMPORT_DATE  As Date

Public BOX As Double                                  'จำนวนต่อกล่อง
Public BOX_AMOUNT As Double            'จำนวนกล่อง
Public TRAY As Double                                 'จำนวนต่อถาด
Public TRAY_AMOUNT As Double           'จำนวนถาด
Public PACK As Double                                 'จำนวนต่อแพ็ค
Public PACK_AMOUNT As Double           'จำนวนแพ็ค
Public BASE_UNIT As Double                      'จำนวนหน่วยล่าง
Public TOTAL_AMOUNT As Double          'จำนวนรวม
'------------------------------------------------------------------------------------------------------------------------------------>
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
      SelectStr1 = "SELECT BVD.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM BALANCE_VERIFY_DATE BVD "
      
   End If
      
   If BALANCE_VERIFY_DATE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BVD.BALANCE_VERIFY_DATE_ID = " & BALANCE_VERIFY_DATE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BVD.BALANCE_VERIFY_DATE_ID = " & BALANCE_VERIFY_DATE_ID & ") "
      End If
   End If
   
   If BALANCE_VERIFY_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BVD.BALANCE_VERIFY_DETAIL_ID = " & BALANCE_VERIFY_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BVD.BALANCE_VERIFY_DETAIL_ID = " & BALANCE_VERIFY_DETAIL_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   HaveOrder = True
   If Ind = 1 Then
      OrderByStr = "ORDER BY BVD.BALANCE_VERIFY_DATE_ID "
      OrderType = 1
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
      BALANCE_VERIFY_DATE_ID = NVLI(Rs("BALANCE_VERIFY_DATE_ID"), -1)
      BALANCE_VERIFY_DETAIL_ID = NVLI(Rs("BALANCE_VERIFY_DETAIL_ID"), -1)
      IMPORT_DATE = InternalDateToDate(NVLS(Rs("IMPORT_DATE"), ""))
      
      BOX = NVLD(Rs("BOX"), 0)
      BOX_AMOUNT = NVLD(Rs("BOX_AMOUNT"), 0)
      TRAY = NVLD(Rs("TRAY"), 0)
      TRAY_AMOUNT = NVLD(Rs("TRAY_AMOUNT"), 0)
      PACK = NVLD(Rs("PACK"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      BASE_UNIT = NVLD(Rs("BASE_UNIT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
   End If
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("BALANCE_VERIFY_DATE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BALANCE_VERIFY_DATE " & _
                     "( " & _
                     "   BALANCE_VERIFY_DETAIL_ID, " & _
                     "   IMPORT_DATE, " & _
                     "   BOX, " & _
                     "   BOX_AMOUNT, " & _
                     "   TRAY, " & _
                     "   TRAY_AMOUNT, " & _
                     "   PACK, " & _
                     "   PACK_AMOUNT, " & _
                     "   BASE_UNIT, " & _
                     "   TOTAL_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     BALANCE_VERIFY_DETAIL_ID & "," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(IMPORT_DATE))) & "'," & _
                     BOX & "," & _
                     BOX_AMOUNT & "," & _
                     TRAY & "," & _
                     TRAY_AMOUNT & "," & _
                     PACK & "," & _
                     PACK_AMOUNT & "," & _
                     BASE_UNIT & "," & _
                     TOTAL_AMOUNT & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         BALANCE_VERIFY_DATE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_VERIFY_DATE SET " & _
                     "   IMPORT_DATE= '" & ChangeQuote(Trim(DateToStringInt(IMPORT_DATE))) & "', " & _
                     "  BOX = " & BOX & ", " & _
                     "  BOX_AMOUNT = " & BOX_AMOUNT & ", " & _
                     "  TRAY = " & TRAY & ", " & _
                     "  TRAY_AMOUNT = " & TRAY_AMOUNT & ", " & _
                     "  PACK = " & PACK & ", " & _
                     "  PACK_AMOUNT = " & PACK_AMOUNT & ", " & _
                     "  BASE_UNIT = " & BASE_UNIT & ", " & _
                     "  TOTAL_AMOUNT = " & TOTAL_AMOUNT & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_VERIFY_DATE_ID = " & BALANCE_VERIFY_DATE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM BALANCE_VERIFY_DATE WHERE BALANCE_VERIFY_DATE_ID = " & BALANCE_VERIFY_DATE_ID
      m_Conn.Execute (SQL1)
End Function
