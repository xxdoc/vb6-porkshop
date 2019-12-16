VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterFromToEx"
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

Public MASTER_FROMTO_EX_ID As Long
Public MASTER_FROMTO_ID As Long

Public CUSTOMER_ID As Long
Public PART_ITEM_ID As Long
Public EMP_ID As Long
Public PERCENT  As Double

'--------------------------------------------------------------------------------------------------------------------- >
Public VALID_FROM As Date
Public VALID_TO As Date
Public MASTER_FROMTO_TYPE  As MASTER_COMMISSION_AREA

Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public PART_NO As String
Public PART_DESC As String
Public SALE_CODE As String
Public SALE_NAME  As String

Public CUSTOMER_ID_SET As String
Public PART_ITEM_ID_SET As String
Public EMP_ID_SET As String

Public MASTER_FROMTO_NO  As String
Public MASTER_FROMTO_DESC As String
Public GROUP_COM_ID As Long
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
Dim TempStr  As String

   If Ind = 1 Then
      SelectStr1 = "SELECT MFE.* "
      SelectStr1 = SelectStr1 & ", NM2.LONG_NAME CUSTOMER_NAME, APM.APAR_CODE CUSTOMER_CODE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr1 = SelectStr1 & ", SC.STOCK_NO PART_NO, SC.STOCK_DESC PART_DESC "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MASTER_FROMTO_EX MFE "
      
      FromStr = FromStr & "LEFT OUTER JOIN APAR_MAS APM ON (MFE.CUSTOMER_ID = APM.APAR_MAS_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN APAR_NAME EN ON (APM.APAR_MAS_ID = EN.APAR_MAS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MFE.EMP_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC ON (MFE.PART_ITEM_ID = SC.STOCK_CODE_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT MFE.*,MFT.* "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE SALE_CODE,NM3.LONG_NAME SALE_LONG_NAME,NM3.LAST_NAME SALE_LAST_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MASTER_FROMTO_EX MFE "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_FROMTO MFT ON (MFE.MASTER_FROMTO_ID = MFT.MASTER_FROMTO_ID ) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (MFE.EMP_ID = EM.EMP_ID ) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EM.EMP_ID = EN1.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN1.NAME_ID = NM3.NAME_ID) "
      
      TempStr = "WHERE "
      TempDate = DateToStringIntLow(VALID_FROM)
      WhereStr = WhereStr & TempStr & " (MFT.VALID_FROM <= '" & ChangeQuote(Trim(TempDate)) & "')"
      TempStr = "AND "
      TempDate = DateToStringIntLow(VALID_TO)
      WhereStr = WhereStr & TempStr & " (MFT.VALID_TO >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If MASTER_FROMTO_EX_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MFE.MASTER_FROMTO_EX_ID = " & MASTER_FROMTO_EX_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MFE.MASTER_FROMTO_EX_ID = " & MASTER_FROMTO_EX_ID & ") "
      End If
   End If
   
   If MASTER_FROMTO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MFE.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MFE.MASTER_FROMTO_ID = " & MASTER_FROMTO_ID & ") "
      End If
   End If
   
   If MASTER_FROMTO_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (MFT.MASTER_FROMTO_TYPE = " & MASTER_FROMTO_TYPE & ") " & vbCrLf
   End If
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY MFE.MASTER_FROMTO_EX_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY MFE.MASTER_FROMTO_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      MASTER_FROMTO_EX_ID = NVLI(Rs("MASTER_FROMTO_EX_ID"), -1)
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      PERCENT = NVLD(Rs("PERCENT"), 0)
      
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_LONG_NAME"), "") & " " & NVLS(Rs("SALE_LAST_NAME"), "")
   ElseIf Ind = 2 Then
      MASTER_FROMTO_ID = NVLI(Rs("MASTER_FROMTO_ID"), -1)
      MASTER_FROMTO_NO = NVLS(Rs("MASTER_FROMTO_NO"), "")
      MASTER_FROMTO_DESC = NVLS(Rs("MASTER_FROMTO_DESC"), "")
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      PERCENT = NVLD(Rs("PERCENT"), 0)
      
      SALE_CODE = NVLS(Rs("SALE_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_LONG_NAME"), "") & " " & NVLS(Rs("SALE_LAST_NAME"), "")
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
   
   TempID1 = MASTER_FROMTO_ID
   If MASTER_FROMTO_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID4 = EMP_ID
   If EMP_ID <= 0 Then
      TempID4 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MASTER_FROMTO_EX_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO MASTER_FROMTO_EX " & _
                     "( " & _
                     "   MASTER_FROMTO_EX_ID, " & _
                     "   MASTER_FROMTO_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   EMP_ID, " & _
                     "   PERCENT, " & _
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
                     TempID4 & "," & _
                     PERCENT & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         MASTER_FROMTO_EX_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MASTER_FROMTO_EX SET " & _
                     "   CUSTOMER_ID = " & TempID2 & "," & _
                     "   PART_ITEM_ID = " & TempID3 & "," & _
                     "   EMP_ID = " & TempID4 & "," & _
                     "  PERCENT = " & PERCENT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE MASTER_FROMTO_EX_ID = " & MASTER_FROMTO_EX_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM MASTER_FROMTO_EX WHERE MASTER_FROMTO_EX_ID = " & MASTER_FROMTO_EX_ID
      m_Conn.Execute (SQL1)
     
End Function
