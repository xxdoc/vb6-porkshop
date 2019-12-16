Attribute VB_Name = "modLoadCapital"
Option Explicit
Public Sub LoadCapitalMovementLocation(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PartItem As Long = -1, Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_ITEM_ID = PartItem
   D.LOCATION_ID = LocationID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(1, Rs)
      
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
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadCapitalMovementLocationDocDate(Optional Cl As Collection = Nothing, Optional FromDate As Date, Optional ToDate As Date, Optional PartItem As Long = -1, Optional LocationID As Long)
On Error GoTo ErrorHandler
Dim D As CCapitalMovement
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCapitalMovement
Dim I As Long
   
   Set D = New CCapitalMovement
   Set Rs = New ADODB.Recordset
   
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.PART_ITEM_ID = PartItem
   D.LOCATION_ID = LocationID
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCapitalMovement
      Call TempData.PopulateFromRS(1, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.LOCATION_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
