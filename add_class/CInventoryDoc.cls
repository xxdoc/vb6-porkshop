VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public ImportExportItems As Collection
Public TransferItems As Collection

' INVENTORY_DOC---------------------------------------------------------------------------------------------------------------------------------
Public INVENTORY_DOC_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_DESC  As String
Public DOCUMENT_TYPE As INVENTORY_DOCTYPE

Public INVENTORY_SUB_TYPE As Long
Public CANCEL_FLAG As String
Public YYYYMM  As String
Public APAR_MAS_ID As Long
' INVENTORY_DOC---------------------------------------------------------------------------------------------------------------------------------
' FOR STOCK_CODE --------------------------------------------------------------------------------------------------------------------------------------------\
Public STOCK_NO As String
' FOR STOCK_CODE --------------------------------------------------------------------------------------------------------------------------------------------\

' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public DOCUMENT_TYPE_SET  As String
Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As Long
Public OLD_CANCEL_FLAG  As String
' FOR SEARCH DATA --------------------------------------------------------------------------------------------------------------------------------------------\
Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set ImportExportItems = New Collection
   Set TransferItems = New Collection
End Sub

Private Sub Class_Terminate()
   Set ImportExportItems = Nothing
   Set TransferItems = Nothing
End Sub
Public Sub DeleteData()
Dim SQL As String
   
   SQL = "DELETE FROM LOT_ITEM WHERE INVENTORY_DOC_ID  = " & INVENTORY_DOC_ID
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM INVENTORY_DOC WHERE INVENTORY_DOC_ID  = " & INVENTORY_DOC_ID
   Call m_ADOConn.Execute(SQL)
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   TempID1 = INVENTORY_SUB_TYPE
   If INVENTORY_SUB_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = APAR_MAS_ID
   If APAR_MAS_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If ShowMode = SHOW_ADD Then
      
      SQL1 = "INSERT INTO INVENTORY_DOC " & _
                     "( " & _
                     "   DOCUMENT_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_DESC, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   INVENTORY_SUB_TYPE, " & _
                     "   CANCEL_FLAG, " & _
                     "   YYYYMM, " & _
                     "   APAR_MAS_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                        "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(DOCUMENT_DESC)) & "'," & _
                        DOCUMENT_TYPE & "," & _
                        TempID1 & "," & _
                        "'" & ChangeQuote(Trim(CANCEL_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        TempID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING INVENTORY_DOC_ID "
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_DOC SET " & _
                     "    DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   DOCUMENT_DATE= '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "    DOCUMENT_DESC = '" & ChangeQuote(Trim(DOCUMENT_DESC)) & "', " & _
                     "    DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "    INVENTORY_SUB_TYPE = " & TempID1 & ", " & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    APAR_MAS_ID = " & TempID2 & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
   End If
   
   If ShowMode = SHOW_ADD Then
      INVENTORY_DOC_ID = m_ADOConn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf ShowMode = SHOW_EDIT Then
      Call m_ADOConn.Execute(SQL1)
   End If
End Function

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim NewStr As String
Dim SubLen As Long
Dim TempDate As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT IVTRD.* "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM INVENTORY_DOC IVTRD "
      
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVTRD.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") " & vbCrLf
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVTRD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
   
   If INVENTORY_SUB_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVTRD.INVENTORY_SUB_TYPE = " & INVENTORY_SUB_TYPE & ") " & vbCrLf
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(IVTRD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (IVTRD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
      End If
   End If
   
   If Len(Trim(DOCUMENT_TYPE_SET)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVTRD.DOCUMENT_TYPE IN " & ChangeQuote(Trim(DOCUMENT_TYPE_SET)) & ")"
   End If
      
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVTRD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(Trim(STOCK_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(STOCK_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " ((SELECT COUNT(*) FROM LOT_ITEM LI LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (LI.INVENTORY_DOC_ID = IVTRD.INVENTORY_DOC_ID) AND (SUBSTR(SC.STOCK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)"
      Else
         WhereStr = WhereStr & TempStr & " (((SELECT COUNT(*) FROM LOT_ITEM LI LEFT OUTER JOIN STOCK_CODE SC ON (LI.PART_ITEM_ID = SC.STOCK_CODE_ID) WHERE (LI.INVENTORY_DOC_ID = IVTRD.INVENTORY_DOC_ID) AND SC.STOCK_NO = '" & ChangeQuote(Trim(STOCK_NO)) & "')) > 0)"
      End If
   End If
      
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(CANCEL_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(IVTRD.CANCEL_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (IVTRD.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   If Ind = 3 Then
      'GroupByStr = " GROUP BY IVTRD.DOCUMENT_DATE,SC.STOCK_NO,SC.STOCK_DESC,MR2.KEY_NAME,IVTRD.DOCUMENT_TYPE "
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY IVTRD.DOCUMENT_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY IVTRD.DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY IVTRD.DOCUMENT_NO "
      End If
   End If
   
   If Len(OrderByStr) > 0 Then
      If ORDER_TYPE = 1 Then
         OrderByStr = OrderByStr & " ASC "
      ElseIf ORDER_TYPE = 2 Then
         OrderByStr = OrderByStr & " DESC "
      Else
         OrderByStr = OrderByStr & " DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
      
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      INVENTORY_SUB_TYPE = NVLI(Rs("INVENTORY_SUB_TYPE"), -1)
      
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      
      APAR_MAS_ID = NVLI(Rs("APAR_MAS_ID"), -1)
   End If
End Sub
Public Function UpdateInventorySubType() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   INVENTORY_SUB_TYPE = " & INVENTORY_SUB_TYPE & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & Val(INVENTORY_DOC_ID)

   Call m_ADOConn.Execute(SQL1)
End Function
