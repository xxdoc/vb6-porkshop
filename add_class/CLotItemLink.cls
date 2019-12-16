VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLotItemLink"
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

Public LOT_ITEM_LINK_ID As Long
Public EXPORT_LOT_ITEM_ID As Long      ' LINK ไปยังสินค้าที่ เอา ออกไปผลิต หรือ ขาย LOT ITEM FK ของการเบิกสินค้าใดๆ
Public IMPORT_LOT_ITEM_ID As Long      ' LINK ไปยังสินค้าที่นำเข้า (เป็นการบอกว่า เบิก ITEM มาจาก LOT ITEM ที่นำเข้าอันไหน)
Public IMPORT_AMOUNT As Double    ' จำนวนตาม LOT_ITEM ที่นำเข้า
Public MAIN_IMPORT_LOT_ITEM_ID As Long      ' LINK ไปยังสินค้าที่นำเข้า ขั้นต้น DOCUMENT_TYPE = 1 ใบรับเข้า,22 ใบเสร็จซื้อสด,11 ใบรับสินค้าซื้อสด

'เช่น 1 นำเข้าด้วย LOT_ID 1 จำนวน 1000 และนำไปผลิต เป็น ID 2 จำนวน 500 และขาย ID 3 จำนวน 200 ซึ่งจะมี
'        EXPORT_LOT_ITEM_ID = 2 และ IMPORT_LOT_ITEM_ID  = 1          IMPORT_AMOUNT =  500
'        EXPORT_LOT_ITEM_ID = 3 และ IMPORT_LOT_ITEM_ID  = 1          IMPORT_AMOUNT =  200

'เช่น 1 นำเข้าด้วย LOT_ID X จำนวน 500 และ โอนเข้าด้วย ID Y จำนวน 500 และนำไปผลิต เป็น ID Z จำนวน 800 โดยใช้ X =  500 ,Y  = 300 ซึ่งจะมี
'        EXPORT_LOT_ITEM_ID = Z และ IMPORT_LOT_ITEM_ID  = X          IMPORT_AMOUNT =  500
'        EXPORT_LOT_ITEM_ID = Z และ IMPORT_LOT_ITEM_ID  = Y          IMPORT_AMOUNT =  300

'--------------------------------------------------------------------------------------------------------------------- >

Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE  As Long
Public INVENTORY_DOC_ID  As Long
Public PART_ITEM_ID  As Long
Public PART_NO As String
Public PART_DESC As String
Public UNIT_ID As Long
Public LOT_ITEM_AMOUNT As Double    ' จำนวนตาม LOT_ITEM ที่นำเข้าตั้งแต่แรกก่อนเลือกผลิต
Public AVG_PRICE As Double

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date
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
Dim TempStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT LK.* "
      SelectStr1 = SelectStr1 & ",IVD1.DOCUMENT_DATE,IVD1.DOCUMENT_NO "
      SelectStr1 = SelectStr1 & ",LT1.PART_ITEM_ID,SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM LOT_ITEM_LINK LK "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (LT1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (LT1.PART_ITEM_ID = SC1.STOCK_CODE_ID) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT LK.* "
      SelectStr1 = SelectStr1 & ",IVD1.INVENTORY_DOC_ID,IVD1.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",LT1.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM LOT_ITEM_LINK LK "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.LOT_ITEM_ID = LK.IMPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (LT1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT LK.EXPORT_LOT_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM LOT_ITEM_LINK LK "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.LOT_ITEM_ID = LK.EXPORT_LOT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (LT1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
   
   
   
   End If
   
   If LOT_ITEM_LINK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LK.LOT_ITEM_LINK_ID = " & LOT_ITEM_LINK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LK.LOT_ITEM_LINK_ID = " & LOT_ITEM_LINK_ID & ") "
      End If
   End If
   
   If EXPORT_LOT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LK.EXPORT_LOT_ITEM_ID = " & EXPORT_LOT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LK.EXPORT_LOT_ITEM_ID = " & EXPORT_LOT_ITEM_ID & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVD1.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (IVD1.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         
      ElseIf OrderBy = 2 Then
         
      ElseIf OrderBy = 3 Then
         
     Else
         OrderByStr = "ORDER BY IVD1.DOCUMENT_DATE,IVD1.DOCUMENT_NO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY LK.LOT_ITEM_LINK_ID "
      OrderType = 1
   ElseIf Ind = 3 Then
      
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   If Ind = 3 Then
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      LOT_ITEM_LINK_ID = NVLI(Rs("LOT_ITEM_LINK_ID"), -1)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_LOT_ITEM_ID = NVLI(Rs("EXPORT_LOT_ITEM_ID"), -1)
      IMPORT_LOT_ITEM_ID = NVLI(Rs("IMPORT_LOT_ITEM_ID"), -1)
      MAIN_IMPORT_LOT_ITEM_ID = NVLI(Rs("MAIN_IMPORT_LOT_ITEM_ID"), -1)
      
      '-------------------------------------------------------------------------------------------------------------------------
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 2 Then
      LOT_ITEM_LINK_ID = NVLI(Rs("LOT_ITEM_LINK_ID"), -1)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_LOT_ITEM_ID = NVLI(Rs("EXPORT_LOT_ITEM_ID"), -1)
      IMPORT_LOT_ITEM_ID = NVLI(Rs("IMPORT_LOT_ITEM_ID"), -1)
      MAIN_IMPORT_LOT_ITEM_ID = NVLI(Rs("MAIN_IMPORT_LOT_ITEM_ID"), -1)
      
      '-------------------------------------------------------------------------------------------------------------------------
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ElseIf Ind = 3 Then
      EXPORT_LOT_ITEM_ID = NVLI(Rs("EXPORT_LOT_ITEM_ID"), -1)
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
   
   TempID1 = EXPORT_LOT_ITEM_ID
   If EXPORT_LOT_ITEM_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = IMPORT_LOT_ITEM_ID
   If IMPORT_LOT_ITEM_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = MAIN_IMPORT_LOT_ITEM_ID
   If MAIN_IMPORT_LOT_ITEM_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      'Call glbDatabaseMngr.GetSeqID("LOT_ITEM_LINK_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO LOT_ITEM_LINK " & _
                     "( " & _
                     "   EXPORT_LOT_ITEM_ID, " & _
                     "   IMPORT_LOT_ITEM_ID, " & _
                     "   IMPORT_AMOUNT, " & _
                     "   MAIN_IMPORT_LOT_ITEM_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     IMPORT_AMOUNT & "," & _
                     TempID3 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
         'LOT_ITEM_LINK_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LOT_ITEM_LINK SET " & _
                     "   EXPORT_LOT_ITEM_ID = " & TempID1 & "," & _
                     "   IMPORT_LOT_ITEM_ID = " & TempID2 & "," & _
                     "  IMPORT_AMOUNT = " & IMPORT_AMOUNT & ", " & _
                     "   MAIN_IMPORT_LOT_ITEM_ID = " & TempID3 & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOT_ITEM_LINK_ID = " & LOT_ITEM_LINK_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM LOT_ITEM_LINK WHERE LOT_ITEM_LINK_ID = " & LOT_ITEM_LINK_ID
      m_Conn.Execute (SQL1)
      
End Function
