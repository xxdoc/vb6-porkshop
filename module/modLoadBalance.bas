Attribute VB_Name = "modLoadBalance"
Option Explicit
Public Sub LoadLeftAmountLocation(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long, Optional FromStockNo As String, Optional ToStockNo As String, Optional PartItem As Long, Optional OUTLAY_FLAG As Long = 1)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.ORDER_BY = 1
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   D.PART_ITEM_ID = PartItem
   Call D.QueryData(32, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(32, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub LoadLeftAmountLotItem(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.ORDER_BY = 1
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   Call D.QueryData(20, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(20, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub LoadLeftAmount(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long

   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.FROM_DOC_DATE = FromDate
   D.TO_DOC_DATE = ToDate
   D.LOCATION_ID = LocationID
   D.ORDER_BY = 1
   D.FROM_STOCK_NO = FromStockNo
   D.TO_STOCK_NO = ToStockNo
   Call D.QueryData(30, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(30, Rs)
         
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PART_ITEM_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Function LoadCheckBalance(CompareUseAmount As Double, LocationID As Long, PartItemID As Long, PartNo As String, Optional ExCludeLotID As Long) As Boolean
On Error GoTo ErrorHandler
Dim D As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim I As Long
   
   LoadCheckBalance = False
   Set D = New CLotItem
   Set Rs = New ADODB.Recordset
   
   D.LOCATION_ID = LocationID
   D.PART_ITEM_ID = PartItemID
   D.LOT_ITEM_ID = ExCludeLotID
   Call D.QueryData(29, Rs, ItemCount, False)
   
   Call D.PopulateFromRS(29, Rs)
         
   If D.SUM_AMOUNT >= CompareUseAmount Then
      LoadCheckBalance = True
   Else
      glbErrorLog.LocalErrorMsg = "มียอด " & PartNo & " ไม่เพียงพอสำหรับเบิก   ( มียอดคงเหลือเพียง " & D.SUM_AMOUNT & " )"
      glbErrorLog.ShowUserError
   End If
   
   Set Rs = Nothing
   Set D = Nothing
   
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Function
Public Sub GetSumMovementPartItemTypeDocDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationID
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(42, Rs, ItemCount, True)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(42, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_TYPE & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub GetSumMovementPartItemType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Lt.LOCATION_ID = LocationID
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(10, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(10, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_TYPE & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

