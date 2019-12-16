VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocItemLink"
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

'<--------------------------------------------------------------------------------------------------------------------->
Public DOC_ITEM_LINK_ID As Long
Public DOC_ITEM_ID As Long      ' LINK ไปยังสินค้าที่ เอา ออกไปผลิต หรือ ขาย DOC ITEM FK ของส่งสินค้า หรือ ใบ ส่งของใดๆ
Public IMPORT_LOT_ITEM_ID As Long      ' LINK ไปยังสินค้าที่นำเข้า (เป็นการบอกว่า เบิก ITEM มาจาก LOT ITEM ที่นำเข้าอันไหน)
Public IMPORT_AMOUNT As Double    ' จำนวนตาม LOT_ITEM ที่นำเข้า
Public MAIN_IMPORT_LOT_ITEM_ID As Long      ' LINK ไปยังสินค้าที่นำเข้า ขั้นต้น DOCUMENT_TYPE = 1 ใบรับเข้า,22 ใบเสร็จซื้อสด,11 ใบรับสินค้าซื้อสด
'โดยหลักๆแล้ว Table นี้มีใว้เพื่อใช้ในการ copy จาก DocItemLink ไปยัง LotItemLink
'--------------------------------------------------------------------------------------------------------------------- >

Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE  As Long
Public INVENTORY_DOC_ID  As Long
Public PART_ITEM_ID  As Long
Public PART_NO As String
Public PART_DESC As String

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
      SelectStr1 = "SELECT DIK.* "
'      SelectStr1 = SelectStr1 & ",IVD1.DOCUMENT_DATE,IVD1.DOCUMENT_NO "
'      SelectStr1 = SelectStr1 & ",LT1.PART_ITEM_ID,SC1.STOCK_NO PART_NO,SC1.STOCK_DESC PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM DOC_ITEM_LINK DIK "
'      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LT1 ON (LT1.LOT_ITEM_ID = DIK.IMPORT_LOT_ITEM_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (LT1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
'
'      FromStr = FromStr & "LEFT OUTER JOIN STOCK_CODE SC1 ON (LT1.PART_ITEM_ID = SC1.STOCK_CODE_ID) "
   End If
   
   If DOC_ITEM_LINK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DIK.DOC_ITEM_LINK_ID = " & DOC_ITEM_LINK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DIK.DOC_ITEM_LINK_ID = " & DOC_ITEM_LINK_ID & ") "
      End If
   End If
   
   If DOC_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DIK.DOC_ITEM_ID = " & DOC_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DIK.DOC_ITEM_ID = " & DOC_ITEM_ID & ") "
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
         'OrderByStr = "ORDER BY IVD1.DOCUMENT_DATE,IVD1.DOCUMENT_NO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY DIK.DOC_ITEM_LINK_ID "
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
   
   If Ind = 1 Then
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      DOC_ITEM_LINK_ID = NVLI(Rs("DOC_ITEM_LINK_ID"), -1)
      DOC_ITEM_ID = NVLI(Rs("DOC_ITEM_ID"), -1)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      IMPORT_LOT_ITEM_ID = NVLI(Rs("IMPORT_LOT_ITEM_ID"), -1)
      MAIN_IMPORT_LOT_ITEM_ID = NVLI(Rs("MAIN_IMPORT_LOT_ITEM_ID"), -1)
      
      '-------------------------------------------------------------------------------------------------------------------------
'      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
'      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
'
'      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
'      PART_NO = NVLS(Rs("PART_NO"), "")
'      PART_DESC = NVLS(Rs("PART_DESC"), "")
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
   
   TempID1 = DOC_ITEM_ID
   If DOC_ITEM_ID <= 0 Then
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
      'Call glbDatabaseMngr.GetSeqID("DOC_ITEM_LINK_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO DOC_ITEM_LINK " & _
                     "( " & _
                     "   DOC_ITEM_ID, " & _
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
         'DOC_ITEM_LINK_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE DOC_ITEM_LINK SET " & _
                     "   DOC_ITEM_ID = " & TempID1 & "," & _
                     "   IMPORT_LOT_ITEM_ID = " & TempID2 & "," & _
                     "  IMPORT_AMOUNT = " & IMPORT_AMOUNT & ", " & _
                     "   MAIN_IMPORT_LOT_ITEM_ID = " & TempID3 & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DOC_ITEM_LINK_ID = " & DOC_ITEM_LINK_ID
   End If
   ''debug.print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM DOC_ITEM_LINK WHERE DOC_ITEM_LINK_ID = " & DOC_ITEM_LINK_ID
      m_Conn.Execute (SQL1)
      
End Function
