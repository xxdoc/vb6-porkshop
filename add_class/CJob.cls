VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJob"
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

Public JOB_ID As Long
Public JOB_NO As String
Public JOB_DESC As String
Public JOB_DATE As Date


Public INVENTORY_DOC_ID As Long           'Link づ学

Public FORMULA_ID As Long
Public FORMULA_AMOUNT As Double     '峦吹亚べ车伊寿得
'--------------------------------------------------------------------------------------------------------------------- >

Public KEY_LOOKUP As String
Public KEY_ID As Long

Public ID As Long

Public FORMULA_NO As String
Public FORMULA_DESC As String

Public FROM_DATE As Date
Public TO_DATE As Date

Public STOCK_NO As String                       'For Search

Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As Long

Public CollJobInputs As Collection
Public CollJobOutputs As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CollJobInputs = New Collection
   Set CollJobOutputs = New Collection
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
Dim TempStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT J.* "
      SelectStr1 = SelectStr1 & ",FM.FORMULA_NO,FM.FORMULA_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN FORMULA FM ON (J.FORMULA_ID = FM.FORMULA_ID) "
      
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
   
   If Trim(JOB_NO) <> "" Then
      If WildCard(JOB_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(J.JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(J.JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (J.JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (J.JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         End If
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(JOB_DESC) <> "" Then
      If WildCard(JOB_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "JOB_DESC", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "JOB_DESC", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (J.JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (J.JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "')"
         End If
      End If
   End If
      
   If Len(Trim(STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(STOCK_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM JOB_ITEM JI LEFT OUTER JOIN STOCK_CODE SC ON (JI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (JI.JOB_ID = J.JOB_ID) AND (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)"
      Else
         WhereStr = WhereStr & TempStr & " (((SELECT COUNT(*) FROM JOB_ITEM JI LEFT OUTER JOIN STOCK_CODE SC ON (JI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (JI.JOB_ID = J.JOB_ID) AND SC.STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "')) > 0)"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PRODUCT_NO, PI.PRODUCT_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY J.JOB_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY J.JOB_DATE "
     Else
         OrderByStr = "ORDER BY J.JOB_ID "
      End If
   End If
   
   If Len(OrderByStr) > 0 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      JOB_DESC = NVLS(Rs("JOB_DESC"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
      FORMULA_AMOUNT = NVLD(Rs("FORMULA_AMOUNT"), 0)
      
      FORMULA_NO = NVLS(Rs("FORMULA_NO"), "")
      FORMULA_DESC = NVLS(Rs("FORMULA_DESC"), "")
      
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

   TempID1 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = FORMULA_ID
   If FORMULA_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO JOB " & _
                     "( " & _
                     "   JOB_NO, " & _
                     "   JOB_DESC, " & _
                     "   JOB_DATE, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   FORMULA_ID, " & _
                     "   FORMULA_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     "'" & ChangeQuote(Trim(JOB_NO)) & "'," & _
                     "'" & ChangeQuote(Trim(JOB_DESC)) & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "'," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     FORMULA_AMOUNT & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ")  RETURNING JOB_ID "
   Else 'Show edit
      SQL1 = "UPDATE JOB SET " & _
                     "   JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "', " & _
                     "  JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "', " & _
                     "   JOB_DATE= '" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "', " & _
                     "   INVENTORY_DOC_ID = " & TempID1 & "," & _
                     "   FORMULA_ID = " & TempID2 & "," & _
                     "  FORMULA_AMOUNT = " & FORMULA_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_ID = " & JOB_ID
   End If
   ''debug.print (SQL1)
   If AddEditMode = SHOW_ADD Then
      JOB_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM JOB_ITEM WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
     
      SQL1 = "DELETE FROM JOB WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM INVENTORY_DOC WHERE   INVENTORY_DOC_ID   = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
      
End Function
Private Sub Class_Terminate()
   Set CollJobInputs = Nothing
   Set CollJobOutputs = Nothing
End Sub
Private Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim I As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For I = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next I
    WhereStr = WhereStr & ")"
    
    GenerateSearchLike = WhereStr
End Function

