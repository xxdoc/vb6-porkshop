Attribute VB_Name = "modLoadData"
Option Explicit

Public Sub LoadMaster(C As ComboBox, Optional Cl As Collection = Nothing, Optional KeyType As Long = -1, Optional ShowType As Long = -1, Optional MasterArea As MASTER_TYPE, Optional ParentID As Long = -1, Optional IndexLink As Long = -1)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = -1
   TempData.MASTER_AREA = MasterArea
   TempData.PARENT_ID = ParentID
   TempData.INDEX_LINK = IndexLink
   Call TempData.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         If ShowType = 2 Then
            C.AddItem (TempData.KEY_NAME & "  (" & TempData.KEY_CODE & ")")
         Else
            C.AddItem (TempData.KEY_NAME)
         End If
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.KEY_ID)))
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
Public Sub LoadMasterID(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterArea As Long = -1)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long
   
   Set Rs = Nothing
   Set Rs = New ADODB.Recordset
   Set TempData = New CMasterRef
   
   TempData.KEY_ID = -1
   TempData.MASTER_AREA = MasterArea
   Call TempData.QueryData(5, Rs, ItemCount, True)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.KEY_CODE)))
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
Public Sub LoadEnterprise(Ep As CEnterprise, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CEnterprise
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEnterprise
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Ep.ENTERPRISE_ID = -1
   Call Ep.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEnterprise
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ENTERPRISE_NAME)
         C.ItemData(I) = TempData.ENTERPRISE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.ENTERPRISE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadUserGroup(Ug As CUserGroup, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserGroup
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserGroup
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Ug.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("GROUP_NAME"))
         C.ItemData(I) = TempData.GetFieldValue("GROUP_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("GROUP_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadEmployee(Optional Emp As CEmployee = Nothing, Optional C As ComboBox = Nothing, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmployee
Dim I As Long
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Rs = New ADODB.Recordset
      
      Emp.EMP_ID = -1
      Emp.ORDER_TYPE = 1
      
      Call Emp.QueryData(1, Rs, ItemCount, False)
      
      While Not Rs.EOF
         I = I + 1
         Set TempData = New CEmployee
         Call TempData.PopulateFromRS(1, Rs)
      
         If Not (C Is Nothing) Then
            C.AddItem (Trim(TempData.EMP_NAME))
            C.ItemData(I) = TempData.EMP_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.EMP_ID)))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
      
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
      Set Rs = Nothing
   ElseIf Not (C Is Nothing) Then
      For Each TempData In m_EmployeeColl
         I = I + 1
         C.AddItem (Trim(TempData.EMP_NAME))
         C.ItemData(I) = TempData.EMP_ID
      Next TempData
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadApArMas(Optional ApAr As CAPARMas = Nothing, Optional C As ComboBox = Nothing, Optional Cl As Collection = Nothing, Optional ApArInd As Long = 1)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPARMas
Dim I As Long
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   If Not (Cl Is Nothing) Then
      Set Rs = New ADODB.Recordset
      
      ApAr.APAR_MAS_ID = -1
      ApAr.ORDER_TYPE = 1
      
      Call ApAr.QueryData(1, Rs, ItemCount, False)
      
      While Not Rs.EOF
         I = I + 1
         Set TempData = New CAPARMas
         Call TempData.PopulateFromRS(1, Rs)
      
         If Not (C Is Nothing) Then
            C.AddItem (TempData.APAR_NAME)
            C.ItemData(I) = TempData.APAR_MAS_ID
         End If
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData, Trim(Str(TempData.APAR_MAS_ID)))
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
      
      If Rs.State = adStateOpen Then
         Rs.Close
      End If
      Set Rs = Nothing
   ElseIf Not (C Is Nothing) Then
      If ApArInd = 1 Then
         For Each TempData In m_CustomerColl
            I = I + 1
            C.AddItem (TempData.APAR_NAME)
            C.ItemData(I) = TempData.APAR_MAS_ID
         Next TempData
      ElseIf ApArInd = 2 Then
         For Each TempData In m_SupplierColl
            I = I + 1
            C.AddItem (TempData.APAR_NAME)
            C.ItemData(I) = TempData.APAR_MAS_ID
         Next TempData
      End If
      
   End If
   
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub



Public Sub InitEnterpriseOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสบริษัท"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อบริษัท"))
   C.ItemData(2) = 2
End Sub

Public Sub InitSupplierOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสผู้ค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อผู้ค้า"))
   C.ItemData(2) = 2
End Sub

Public Sub InitCustomerOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub
Public Sub InitMasterOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัส"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("รายละเอียด"))
   C.ItemData(2) = 2
End Sub

Public Sub InitJournalOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitUserOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ชื่อผู้ใช้"))
   C.ItemData(1) = 1
End Sub

Public Sub InitEmployeeOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อพนักงาน"))
   C.ItemData(2) = 2
End Sub

Public Sub InitInventoryDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 1
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub

Public Sub InitBillingDocOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = -1
   
   C.AddItem (MapText("เลขที่เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เอกสาร"))
   C.ItemData(2) = 2
End Sub
Public Sub InitPartItemOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสสต็อค"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อรายการสต็อค"))
   C.ItemData(2) = 2
End Sub

Public Sub InitEmptyCombo(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
End Sub

Public Sub InitOrderType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("น้อยไปมาก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("มากไปน้อย"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReportUnitType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หน่วยหลัก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("หน่วยย่อย"))
   C.ItemData(2) = 2
End Sub
Public Sub InitAdjustType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปรับลด"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ปรับเพิ่ม"))
   C.ItemData(2) = 2
End Sub
Public Sub InitTransportDetailOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("คนขับ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ทะเบียน"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("ขนส่ง"))
   C.ItemData(3) = 3
End Sub

Public Sub InitUserGroupOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(1) = 1
End Sub

Public Sub InitUserStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใช้งานได้")
   C.ItemData(1) = 1

   C.AddItem ("ถูกระงับ")
   C.ItemData(2) = 2
End Sub

Public Sub InitAPAR(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ลูกค้า")
   C.ItemData(1) = 1

   C.AddItem ("ผู้ค้า")
   C.ItemData(2) = 2
End Sub
Public Sub InitShortCode(C As ComboBox)
   C.Clear
   
   C.AddItem ("เฉพาะห้างสรรพสินค้า")
   C.ItemData(0) = 1
   
   C.AddItem ("ไม่รวมห้างสรรพสินค้า")
   C.ItemData(1) = 2

   C.AddItem ("แสดงทั้งหมด")
   C.ItemData(2) = 0
End Sub

Public Sub InitDrCr(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("เดบิต")
   C.ItemData(1) = 1

   C.AddItem ("เครดิต")
   C.ItemData(2) = 2
End Sub

Public Sub InitSellType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("สินค้า")
   C.ItemData(1) = 1

   C.AddItem ("บริการ")
   C.ItemData(2) = 2

   C.AddItem ("กำหนดเอง")
   C.ItemData(3) = 3
End Sub
Public Sub LoadAccessRight(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGroupRight
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupRight
Dim I As Long

   Set D = New CGroupRight
   Set Rs = New ADODB.Recordset
   
   Call D.SetFieldValue("GROUP_RIGHT_ID", -1)
   Call D.SetFieldValue("GROUP_ID", GroupID)
   Call D.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGroupRight
      Call TempData.PopulateFromRS(3, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("RIGHT_ITEM_NAME"))
         C.ItemData(I) = TempData.GetFieldValue("GROUP_RIGHT_ID")
      End If
      
      'Debug.Print TempData.GetFieldValue("RIGHT_ID") & "-" & TempData.GetFieldValue("RIGHT_ITEM_NAME")
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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

Public Sub LoadStockCode(C As ComboBox, Optional Cl As Collection = Nothing, Optional StockType As Long = -1)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCode
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCode
   
   TempData.STOCK_CODE_ID = -1
   TempData.STOCK_TYPE = StockType
   TempData.EXCEPTION_FLAG = "N"
   Call TempData.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockCode
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STOCK_DESC)
         C.ItemData(I) = TempData.STOCK_CODE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.STOCK_CODE_ID)))
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
Public Sub LoadStockCodeFromTo(C As ComboBox, Optional Cl As Collection = Nothing, Optional FromStock As String, Optional ToStock As String, Optional TempType As Long = 0)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CStockCode
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set TempData = New CStockCode
   
   TempData.FROM_STOCK_NO = FromStock
   TempData.TO_STOCK_NO = ToStock
   Call TempData.QueryData(7, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CStockCode
      Call TempData.PopulateFromRS(7, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.STOCK_DESC)
         C.ItemData(I) = TempData.STOCK_CODE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         If TempType = 1 Then
            Call Cl.add(TempData, Trim(TempData.STOCK_NO))
         Else
            Call Cl.add(TempData)
         End If
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
Public Sub LoadEnterpriseAddress(Ad As CAddress, C As ComboBox, Optional Cl As Collection = Nothing, Optional ShowFirst As Boolean = True)
Dim D As CAddress
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
Dim TempIndex As Long

   TempIndex = 0
   Set Rs = New ADODB.Recordset
   
   Call Ad.SetFieldValue("ADDRESS_ID", -1)
   Call Ad.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.GetFieldValue("ADDRESS_ID")
      End If
      If (I > 0) And ShowFirst Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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

Public Sub LoadAparMasAddress(Ad As CAddress, C As ComboBox, Optional Cl As Collection = Nothing, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long
Dim TempIndex As Long
Dim Mask As Long
   TempIndex = 0
   Mask = 0
   Set Rs = New ADODB.Recordset
   
   Call Ad.SetFieldValue("ADDRESS_ID", -1)
   Call Ad.QueryData(3, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.GetFieldValue("ADDRESS_ID")
         If TempData.GetFieldValue("MAIN_FLAG") = "Y" Then
            Mask = I
            C.ListIndex = I
         End If
      End If
      If (I > 0) And ShowFirst And Mask = 0 And Not (C Is Nothing) Then
         C.ListIndex = 1
      End If
      
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("ADDRESS_ID"))))
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

Public Sub InitReportChequeBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่เช็ค"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่เช็ค"))
   C.ItemData(2) = 2
End Sub

Public Sub LoadUpdateRcpCnDn(Rcp As CRcpCnDn_Item, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CRcpCnDn_Item
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRcpCnDn_Item
Dim I As Long

   Set Rs = New ADODB.Recordset

   
   Call Rcp.QueryData(2, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CRcpCnDn_Item
      Call TempData.PopulateFromRS(2, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PAID_AMOUNT"))
         C.ItemData(I) = TempData.GetFieldValue("DOC_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("DOC_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitReportS_1_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("คนขับ และ ขนส่ง"))
   C.ItemData(1) = 2
End Sub
Public Sub InitReportS_2_1Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("สาขา"))
   C.ItemData(1) = 1
      
   C.AddItem (MapText("ยอดขาย"))
   C.ItemData(2) = 2

End Sub
Public Sub InitReport6_2Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("รหัสสินค้า/วัตถุดิบ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ชื่อสินค้า/วัตถุดิบ"))
   C.ItemData(2) = 2
   
End Sub
Public Sub InitSaleType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("พนักงานขายคิดจากราคาขายเป็นหลัก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("พนักงานขายคิดจากจำนวนขายเป็นหลัก"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("หัวหน้าพนักงานขาย"))
   C.ItemData(3) = 3
   
   C.AddItem (MapText("พนักงานขายคิดจากเฉพาะจำนวนขาย"))
   C.ItemData(4) = 4
   
End Sub
Public Function IdToSaleType(ID As Long) As String
   If ID = 1 Then
      IdToSaleType = "พนักงานขายคิดจากราคาขายเป็นหลัก"
   ElseIf ID = 2 Then
      IdToSaleType = "พนักงานขายคิดจากจำนวนขายเป็นหลัก"
   ElseIf ID = 3 Then
      IdToSaleType = "หัวหน้าพนักงานขาย"
   ElseIf ID = 4 Then
      IdToSaleType = "พนักงานขายคิดจากเฉพาะจำนวนขาย"
   End If
End Function
Public Sub InitPackageOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("รหัสการตั้งราคา"))
   C.ItemData(1) = 1
      
   C.AddItem (MapText("รายละเอียด"))
   C.ItemData(2) = 2

End Sub
Public Sub LoadPackage(Package As CPackage, C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterFlag As String = "N")
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackage
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call Package.SetFieldValue("PACKAGE_ID", -1)
   Call Package.SetFieldValue("PACKAGE_MASTER_FLAG", MasterFlag)
   Call Package.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackage
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PACKAGE_NO") & " " & TempData.GetFieldValue("PACKAGE_DESC"))
         C.ItemData(I) = TempData.GetFieldValue("PACKAGE_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("PACKAGE_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadPackageDetail(PackageDetail As CPackageDetail, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CPackageDetail
Dim I As Long

   Set Rs = New ADODB.Recordset
   
   Call PackageDetail.SetFieldValue("PACKAGE_DETAIL_ID", -1)
   Call PackageDetail.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CPackageDetail
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("PACKAGE_DETAIL_ID"))
         C.ItemData(I) = TempData.GetFieldValue("PACKAGE_DETAIL_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitReportS_2_16Orderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
      
   C.AddItem (MapText("ตามปริมาณขาย"))
   C.ItemData(1) = 1
      
   C.AddItem (MapText("ตามยอดขาย"))
   C.ItemData(2) = 2

End Sub
Public Sub InitReportNullOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
End Sub

Public Sub InitThaiMonth(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มกราคม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("กุมภาพันธ์"))
   C.ItemData(2) = 2

C.AddItem (MapText("มีนาคม"))
   C.ItemData(3) = 3

C.AddItem (MapText("เมษายน"))
   C.ItemData(4) = 4

C.AddItem (MapText("พฤษภาคม"))
   C.ItemData(5) = 5

C.AddItem (MapText("มิถุนายน"))
   C.ItemData(6) = 6

C.AddItem (MapText("กรกฎาคม"))
   C.ItemData(7) = 7

C.AddItem (MapText("สิงหาคม"))
   C.ItemData(8) = 8

C.AddItem (MapText("กันยายน"))
   C.ItemData(9) = 9

C.AddItem (MapText("ตุลาคม"))
   C.ItemData(10) = 10

C.AddItem (MapText("พฤษศจิกายน"))
   C.ItemData(11) = 11

   C.AddItem (MapText(" ธันวาคม"))
   C.ItemData(12) = 12
End Sub

Public Sub LoadDocItem(Mr As CDocItem, C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CDocItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CDocItem
Dim I As Long

   Set Rs = New ADODB.Recordset

   Call Mr.SetFieldValue("DOC_ITEM_ID", -1)
   Call Mr.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CDocItem
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GetFieldValue("DOC_ITEM_ID"))
         C.ItemData(I) = TempData.GetFieldValue("DOC_ITEM_ID")
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("DOC_ITEM_ID"))))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitTagetOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เดือนปี"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("ชื่อพนักงาน"))
   C.ItemData(3) = 3
End Sub
Public Sub InitTagetJobOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เดือนปี"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสวัตถุดิบ"))
   C.ItemData(2) = 2
   
   C.AddItem (MapText("รายละเอียดวัตถุดิบ"))
   C.ItemData(3) = 3
End Sub
Public Sub InitImportType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("นำเข้าลูกหนี้ยกมา")
   C.ItemData(1) = 1
   
'   C.AddItem ("นำเข้าสต็อคยกมา")
'   C.ItemData(2) = 2
'
'   C.AddItem ("นำเข้าข้อมูลลูกหนี้")
'   C.ItemData(3) = 3
'
'   C.AddItem ("นำเข้าข้อมูล BALANCE ACCUM")
'   C.ItemData(4) = 4
End Sub
Public Sub LoadConfigDoc(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CConfigDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CConfigDoc
Dim I As Long

   Set D = New CConfigDoc
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CConfigDoc
      Call TempData.PopulateFromRS(1, Rs)
      
      TempData.Flag = "I"
      
      If Not (C Is Nothing) Then
         I = I + 1
         C.AddItem (TempData.GetFieldValue("CONFIG_DOC_CODE"))
         C.ItemData(I) = TempData.GetFieldValue("CONFIG_DOC_TYPE")
      End If

      If Not (Cl Is Nothing) Then
         ''debug.print (Trim(Str(TempData.GetFieldValue("CONFIG_DOC_TYPE"))))
         Call Cl.add(TempData, Trim(Str(TempData.GetFieldValue("CONFIG_DOC_TYPE"))))
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
Public Sub InitDocItemType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("ปกติ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("ไม่แสดง"))
   C.ItemData(2) = 2
   
'   C.AddItem (MapText("ไม่แสดง"))
'   C.ItemData(3) = 3
End Sub
Public Sub LoadSaleAmount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   Call BD.QueryData(12, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(12, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_MAS_ID & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetDistinctBranchEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(40, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(41, Rs1, ItemCount)

   
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 40, Rs1, 41, 4)
      
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctBranchEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(43, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
      
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(44, Rs1, ItemCount)

   
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 43, Rs1, 44, 5)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long


   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(47, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(47, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
      
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub

Public Sub GetSaleAmountStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(39, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(39, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetDistinctBranchAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(48, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(49, Rs1, ItemCount)
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 48, Rs1, 49, 6)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctBranchStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(52, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(53, Rs1, ItemCount)
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 52, Rs1, 53, 7)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(56, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(57, Rs1, ItemCount)

   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 56, Rs1, 57, 3)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   Call BD.QueryData(147, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate

   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   Call BD.QueryData(148, Rs1, ItemCount)

   Set BD = Nothing
   
   Call GetDataToRsPopulate2(Cl, Rs, 147, Rs1, 148, 16)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDebtorAging(Cl As Collection, Cl2 As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional ShortCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset

'   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.BILLING_DOC_ID = -1
   BD.APAR_IND = 1
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   If ShortCode = 0 Then
      ShortCode = ""
   End If
   BD.SHORT_CODE = ShortCode
   Call BD.QueryData(145, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

'   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.BILLING_DOC_ID = -1
   BD.APAR_IND = 1
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.SHORT_CODE = ShortCode
   Call BD.QueryData(146, Rs1, ItemCount)

   Set BD = Nothing

   If Rs.RecordCount = 0 Then
      While Not Rs1.EOF
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS(146, Rs1)
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData)
         End If
         
         Set TempData = Nothing
         Rs1.MoveNext
      Wend
   ElseIf Rs1.RecordCount = 0 Then
      While Not Rs.EOF
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS(145, Rs)
         
         If Not (Cl Is Nothing) Then
            Call Cl.add(TempData)
         End If
         
         Set TempData = Nothing
         Rs.MoveNext
      Wend
   Else
      Call GetDataToRs(Cl, Rs, 145, Rs1, 146, 15)
   End If
   'สำหรับตรวจสอบ
   Set Rs = Nothing
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.BILLING_DOC_ID = -1
'   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.APAR_IND = 1
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.SHORT_CODE = ShortCode
    Call BD.QueryData(13, Rs, ItemCount)
   
   If Not (Cl2 Is Nothing) Then
      Set Cl2 = Nothing
      Set Cl2 = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl2 Is Nothing) Then
         Call Cl2.add(TempData, Trim(TempData.DOCUMENT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeFreeNotBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(135, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(136, Rs1, ItemCount)

   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 135, Rs1, 136, 11)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpGroupStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   BD.ORDER_BY = OrderBy
   Call BD.QueryData(151, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(152, Rs1, ItemCount)

   Set BD = Nothing
   
   Call GetDataToRsPopulate2(Cl, Rs, 151, Rs1, 152, 17)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub LoadNote(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String = "", Optional ToAparCode As String = "", Optional FromSaleCode As String = "", Optional ToSaleCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim TempData As CBillingDoc
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(130, Rs, ItemCount)
   
   Set BD = Nothing
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(144, Rs1, ItemCount)
   
   Set BD = Nothing
   Call GetDataToRs(Cl, Rs, 130, Rs1, 144, 14)

   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpAparDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional OrderBy As Long = 1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(124, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   Call BD.QueryData(128, Rs1, ItemCount)

   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 124, Rs1, 128, 10)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeGroupName(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(118, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(119, Rs1, ItemCount)

   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 118, Rs1, 119, 8)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctEmpAparStockCodeDocTypeGroupName2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(120, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(121, Rs1, ItemCount)

   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 120, Rs1, 121, 9)
   
   

   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(58, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(58, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(59, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(59, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(126, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(126, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_TYPE_CODE & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeYYYYMM3(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(142, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(142, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(127, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(127, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_TYPE_CODE & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExYYYYMM3(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(143, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(143, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeDateFreeNotBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(137, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(137, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
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
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional ConsignmentFlag As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(153, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(153, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeDateFree2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(122, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(122, Rs)
      
'      TempSum = TempSum + TempData.TOTAL_PRICE
'      TempSum2 = TempSum2 + TempData.DISCOUNT_AMOUNT
'      TempSum3 = TempSum3 + TempData.EXT_DISCOUNT_AMOUNT
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   ''debug.print (TempSum & "------" & TempSum2 & "----------" & TempSum3)
   
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountEmpDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_BY = 9999
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(150, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(150, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFreeNotBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(138, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(138, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
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
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeExDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional ConsignmentFlag As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(154, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(154, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
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
Public Sub GetSaleAmountEmpAparStockCodeDocTypeFreeExDateFree2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
      
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(123, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(123, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
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
Public Sub GetDistinctEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   
   Call BD.QueryData(60, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(61, Rs1, ItemCount)
   
   
   Set BD = Nothing
   
   Call GetDataToRs(Cl, Rs, 60, Rs1, 61, 1)
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDataToRs(Cl As Collection, Rs As ADODB.Recordset, Ind1 As Long, Rs1 As ADODB.Recordset, Ind2 As Long, KeyType As Long)
Dim Af As Boolean
Dim Bf As Boolean
Dim TempData As CBillingDoc
Dim TempData1 As CBillingDoc
Dim I As Long
Dim j As Long
   Af = True
   Bf = True
    I = 1
    j = 1
    
   While Not (Rs.EOF And Rs1.EOF)
      If Af Then
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS(Ind1, Rs)
      End If
      If Bf Then
         Set TempData1 = New CBillingDoc
         Call TempData1.PopulateFromRS(Ind2, Rs1)
      End If

      Af = False
      Bf = False
      If Rs.RecordCount >= I And Rs1.RecordCount < j Then
         Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
         Set TempData = Nothing
         If Not (Rs.EOF) Then
            Rs.MoveNext
            Af = True
            I = I + 1
         End If
      ElseIf Rs.RecordCount < I And Rs1.RecordCount >= j Then
         Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
         Set TempData1 = Nothing
         If Not (Rs1.EOF) Then
            Rs1.MoveNext
            Bf = True
            j = j + 1
         End If
      Else
         If GetDataKeyRs(TempData, KeyType) < GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            Set TempData = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) > GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
            Set TempData1 = Nothing
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) = GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            Set TempData = Nothing
            Set TempData1 = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         End If
      End If
   Wend

End Sub
Public Sub GetDataToRsPopulate2(Cl As Collection, Rs As ADODB.Recordset, Ind1 As Long, Rs1 As ADODB.Recordset, Ind2 As Long, KeyType As Long)
Dim Af As Boolean
Dim Bf As Boolean
Dim TempData As CBillingDoc
Dim TempData1 As CBillingDoc
Dim I As Long
Dim j As Long
   Af = True
   Bf = True
    I = 1
    j = 1
    
   While Not (Rs.EOF And Rs1.EOF)
      If Af Then
         Set TempData = New CBillingDoc
         Call TempData.PopulateFromRS2(Ind1, Rs)
      End If
      If Bf Then
         Set TempData1 = New CBillingDoc
         Call TempData1.PopulateFromRS2(Ind2, Rs1)
      End If
      Af = False
      Bf = False
      If Rs.RecordCount >= I And Rs1.RecordCount < j Then
         Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
         Set TempData = Nothing
         If Not (Rs.EOF) Then
            Rs.MoveNext
            Af = True
            I = I + 1
         End If
      ElseIf Rs.RecordCount < I And Rs1.RecordCount >= j Then
         Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
         Set TempData1 = Nothing
         If Not (Rs1.EOF) Then
            Rs1.MoveNext
            Bf = True
            j = j + 1
         End If
      Else
         If GetDataKeyRs(TempData, KeyType) < GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            Set TempData = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) > GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData1, GetDataKeyRs(TempData1, KeyType))
            Set TempData1 = Nothing
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         ElseIf GetDataKeyRs(TempData, KeyType) = GetDataKeyRs(TempData1, KeyType) Then
            Call Cl.add(TempData, GetDataKeyRs(TempData, KeyType))
            Set TempData = Nothing
            Set TempData1 = Nothing
            If Not (Rs.EOF) Then
               Rs.MoveNext
               Af = True
               I = I + 1
            End If
            If Not (Rs1.EOF) Then
               Rs1.MoveNext
               Bf = True
               j = j + 1
            End If
         End If
      End If
   Wend

End Sub
Public Function GetDataKeyRs(DataGet As Object, KeyType As Long) As String
   If KeyType = 1 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 2 Then
      GetDataKeyRs = Trim(DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 3 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.STOCK_NO)
'      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 4 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 5 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.SALE_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 6 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 7 Then
      GetDataKeyRs = Trim(DataGet.CUSTOMER_BRANCH_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 8 Then
      GetDataKeyRs = Trim(DataGet.APAR_GROUP_NAME & "-" & DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 9 Then
      GetDataKeyRs = Trim(DataGet.APAR_GROUP_NAME & "-" & DataGet.SALE_CODE & "-" & DataGet.STOCK_GROUP_NAME & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 10 Then
      GetDataKeyRs = Trim(DataGet.APAR_GROUP_NAME & "-" & DataGet.APAR_CODE)
   ElseIf KeyType = 11 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.APAR_CODE & "-" & DataGet.STOCK_NO)
   ElseIf KeyType = 12 Then
      GetDataKeyRs = Trim(DataGet.STOCK_TYPE_CODE)
   ElseIf KeyType = 13 Then
      GetDataKeyRs = Trim(DataGet.STOCK_NO)
   ElseIf KeyType = 14 Then
      GetDataKeyRs = Trim(DataGet.DOCUMENT_NO & "-" & DataGet.DOCUMENT_DATE)
   ElseIf KeyType = 15 Then
'      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.DOC_ID_BILLS_NO & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.DOCUMENT_NO & "-" & DataGet.DUE_DATE)
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.CUSTOMER_BRANCH & "-" & DataGet.DOCUMENT_NO & "-" & DataGet.Due_Date)
   ElseIf KeyType = 16 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE)
   ElseIf KeyType = 17 Then
      GetDataKeyRs = Trim(DataGet.SALE_CODE & "-" & DataGet.STOCK_TYPE_CODE & "-" & DataGet.STOCK_NO)
   Else
   End If
End Function
Public Sub GetSaleAmountEmpStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode

   Call BD.QueryData(62, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(62, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode

   Call BD.QueryData(63, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(63, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetDistinctAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(64, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
         
   Set BD = Nothing
   
   'ต้อง merge กับ ใน ส่วนของ สาขาย่อย
   
   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(65, Rs1, ItemCount)
   
   Set BD = Nothing
   
   
   Call GetDataToRs(Cl, Rs, 64, Rs1, 65, 2)
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountAparStockCodeDocTypeFreeEx(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(66, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(66, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountAparStockCodeDocTypeFreeEx1(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromSaleCode As String, Optional ToSaleCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   Call BD.QueryData(67, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(67, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetDistinctBillingAddition(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingAddition
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingAddition
Dim I As Long
   
   Set BD = New CBillingAddition
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(2, Rs, ItemCount)
   
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
      Set TempData = New CBillingAddition
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetSumBillingAdditionID(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingAddition
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingAddition
Dim I As Long
   
   Set BD = New CBillingAddition
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(3, Rs, ItemCount)
   
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
      Set TempData = New CBillingAddition
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BILLING_DOC_ID & "-" & TempData.ADDITION_ID))
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
Public Sub GetDistinctBillingSubTract(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CBillingSubTract
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingSubTract
Dim I As Long
   
   Set BD = New CBillingSubTract
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(2, Rs, ItemCount)
   
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
      Set TempData = New CBillingSubTract
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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

Public Sub GetDistinctTransferPartItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationID As Long = -1, Optional ToLocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.FROM_LOCATION_ID = FromLocationID
   Lt.TO_LOCATION_ID = ToLocationID
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(13, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(13, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetDistinctTransferPartItemConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationCode As String = "", Optional FromLocationCode2 As String = "", Optional ToLocationCode As String = "", Optional ToLocationCode2 As String = "", Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional Consign As Long = -1)
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
   Lt.FROM_LOCATION_CODE = FromLocationCode
   Lt.TO_LOCATION_CODE = ToLocationCode
   Lt.FROM_LOCATION_CODE2 = FromLocationCode2
   Lt.TO_LOCATION_CODE2 = ToLocationCode2
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(48, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(48, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetSumTransferPartItemDocDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationID As Long = -1, Optional ToLocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.FROM_LOCATION_ID = FromLocationID
   Lt.TO_LOCATION_ID = ToLocationID
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(12, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(12, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
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
Public Sub GetTransferPartItemDocDateConsignment(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromLocationCode As String = "", Optional FromLocationCode2 As String = "", Optional ToLocationCode As String = "", Optional ToLocationCode2 As String = "", Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional Consign As Long = -1, Optional OUTLAY_FLAG As Long = 1)
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
   Lt.FROM_LOCATION_CODE = FromLocationCode
   Lt.TO_LOCATION_CODE = ToLocationCode
   Lt.FROM_LOCATION_CODE2 = FromLocationCode2
   Lt.TO_LOCATION_CODE2 = ToLocationCode2
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(49, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(49, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID & "-" & TempData.DOCUMENT_NO))
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

Public Sub GetSumBillingSubTractID(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim Bs As CBillingSubTract
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingSubTract
Dim I As Long
   
   
   Set Bs = New CBillingSubTract
   Set Rs = New ADODB.Recordset
   
   Bs.FROM_DATE = FromDate
   Bs.TO_DATE = ToDate
   Call Bs.QueryDataReport(3, Rs, ItemCount)
   
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
      Set TempData = New CBillingSubTract
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BILLING_DOC_ID & "-" & TempData.SUBTRACT_ID))
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
Public Function CashDocTypeToText(Ct As CASH_DOC_TYPE) As String
   If Ct = CHEQUE_REV Then
      CashDocTypeToText = "เช็ครับ"
   ElseIf Ct = CHEQUE_PAY Then
      CashDocTypeToText = "เช็คจ่าย"
   ElseIf Ct = CASH_DEPOSIT Then
      CashDocTypeToText = "ใบนำฝาก"
   ElseIf Ct = POST_CHEQUE Then
      CashDocTypeToText = "ใบ POST เช็ค"
   End If
End Function
Public Sub InitJobOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ JOB"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่ JOB"))
   C.ItemData(2) = 2
End Sub
Public Sub InitBalanceVerifyOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("หมายเลข"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่"))
   C.ItemData(2) = 2
      
End Sub

Public Sub InitFormulaOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ สูตร"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("วันที่ สูตร"))
   C.ItemData(2) = 2
   
End Sub
Public Sub GetBalanceVerifyByDateLocationPartItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional LocationID As Long = -1)
On Error GoTo ErrorHandler
Dim Ji As CBalanceVerifyDeTail
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBalanceVerifyDeTail
Dim I As Long
   
   Set Ji = New CBalanceVerifyDeTail
   Set Rs = New ADODB.Recordset
   
   Ji.FROM_DATE = FromDate
   Ji.TO_DATE = ToDate
   Ji.FROM_STOCK_NO = FromStockNo
   Ji.TO_STOCK_NO = ToStockNo
   Ji.LOCATION_ID = LocationID
   Call Ji.QueryData(2, Rs, ItemCount)
   
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
      Set TempData = New CBalanceVerifyDeTail
      Call TempData.PopulateFromRS(2, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BALANCE_VERIFY_DATE & "-" & TempData.PART_ITEM_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
      
   Set Ji = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetDistinctPartItemByDocTypeDocSubType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(14, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(14, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetDistinctPartItemByProduction(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional TxType = "")
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If TxType = "I" Then
      Call Lt.QueryData(36, Rs, ItemCount, False)
   ElseIf TxType = "E" Then
      'ยังไม่มีใช้
   End If
   
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
      If TxType = "I" Then
         Call TempData.PopulateFromRS(36, Rs)
      Else
         Call TempData.PopulateFromRS(36, Rs)
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetSumAmountByPartItemDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(15, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(15, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
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
Public Sub GetSumAmountByPartItemDateProduction(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional TxType = "")
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   If TxType = "I" Then
      Call Lt.QueryData(37, Rs, ItemCount, False)
   ElseIf TxType = "E" Then
      'ยังไม่มีใช้
   End If
   
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
      If TxType = "I" Then
         Call TempData.PopulateFromRS(37, Rs)
      Else
      
      End If
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.PART_ITEM_ID))
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
Public Sub GetSumAmountByPartItemDateIndSub(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional InventorySubType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.INVENTORY_SUB_TYPE = InventorySubType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(18, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(18, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_DATE & "-" & TempData.INVENTORY_SUB_TYPE & "-" & TempData.PART_ITEM_ID))
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
Public Sub InitInventoryDocType(C As ComboBox)
Dim I As Long

   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0

   For I = IMPORT_DOCTYPE To ADJUST_DOCTYPE
      C.AddItem (Doctype2Text(I))
      C.ItemData(I) = I
   Next I
   
End Sub
Public Sub GetSaleAmountStockCodeDocType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockCode As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockCode
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(77, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(77, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctStockCode_Y(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(165, Rs, ItemCount) '  '80
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(165, Rs)
      'Call TempData.PopulateFromRS(80, Rs)
            
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub

Public Sub GetDistinctStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1, Optional CONSIGNMENT As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(80, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(80, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountStockCodeDocTypeDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(85, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(85, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountAparTypeStockCodeDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(84, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(84, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_TYPE_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountAparGroupStockCodeDateFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(91, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(91, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_DATE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetDistinctBillBankAccount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
      
   
   
   Set BD = New CCashTran
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(3, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(3, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetTransferAmountBillBankAccount(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date)
On Error GoTo ErrorHandler
Dim BD As CCashTran
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCashTran
Dim I As Long
      
   
   
   Set BD = New CCashTran
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   Call BD.QueryDataReport(4, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CCashTran
      Call TempData.PopulateFromRS(4, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.BILLING_DOC_ID & "-" & TempData.BANK_ACCOUNT))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetBalanceByLotItemLinkDate(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Call Lt.QueryData(24, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(24, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetDocItemIDLinkLotItemID(Cl As Collection, Optional LocationID As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String, Optional PartItem As Long, Optional ChkStd As String, Optional FromDate As Date = -1, Optional ToDate As Date = -1)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim I As Long
   
   Set Lt = New CLotItem
   Set Rs = New ADODB.Recordset
   
   Lt.LOCATION_ID = LocationID
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Lt.PART_ITEM_ID = PartItem
   Lt.FROM_DOC_DATE = FromDate
   Lt.TO_DOC_DATE = ToDate
   Call Lt.QueryData(23, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CLotItem
      Call TempData.PopulateFromRS(23, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.LOT_ITEM_ID)))
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
Public Sub GetSaleAmountMonthByStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   
   Call BD.QueryData(24, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(24, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.YYYYMM & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountMonth(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT2_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   Call BD.QueryData(21, Rs, ItemCount)
      
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(21, Rs)

      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_MAS_ID & "-" & TempData.YYYYMM & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountMonthByCustomerStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   Call BD.QueryData(1007, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1007, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.YYYYMM & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional UnitType As Long = -1, Optional TotalSalePrice As Double, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim CreditStockCode As CCreditStockCode
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   Call BD.QueryData(25, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(25, Rs)
      
      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.STOCK_NO), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.STOCK_DESC = TempData.STOCK_DESC
            CreditStockCode.UNIT_NAME = TempData.UNIT_NAME
            CreditStockCode.UNIT_CHANGE_NAME = TempData.UNIT_CHANGE_NAME
            CreditStockCode.UNIT_AMOUNT = TempData.UNIT_AMOUNT
            
            Call Cl.add(CreditStockCode, Trim(TempData.STOCK_NO))
         End If
         CreditStockCode.STOCK_NO = TempData.STOCK_NO
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT
            
            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE
            
            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT
   
         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT
            
            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT
            
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE
            
            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE
            
            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT
            
         End If
      End If
      
      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCode2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional UnitType As Long = -1, Optional TotalSalePrice As Double, Optional FromStockNo As String, Optional ToStockNo As String, Optional InCludeFree As Long = -1, Optional SaleCode As String = "", Optional FromAparCode As String = "", Optional ToAparCode As String = "")
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim TempData As CBillingDoc
Dim CreditStockCode As CCreditStockCode
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.SALE_CODE = SaleCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   Call BD.QueryData(131, Rs, ItemCount)

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   Set BD = Nothing

   Set BD = New CBillingDoc
   Set Rs1 = New ADODB.Recordset

   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.SALE_CODE = SaleCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   Call BD.QueryData(132, Rs1, ItemCount)

   
   Set BD = Nothing

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(131, Rs)

      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.STOCK_NO), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.STOCK_DESC = TempData.STOCK_DESC
            CreditStockCode.UNIT_NAME = TempData.UNIT_NAME
            CreditStockCode.UNIT_CHANGE_NAME = TempData.UNIT_CHANGE_NAME
            CreditStockCode.UNIT_AMOUNT = TempData.UNIT_AMOUNT

            Call Cl.add(CreditStockCode, Trim(TempData.STOCK_NO))
         End If
         CreditStockCode.STOCK_NO = TempData.STOCK_NO
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT

            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT

         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT

            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE

            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT

         End If
      End If

      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs.MoveNext
   Wend

   While Not Rs1.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(132, Rs1)

      If Not (Cl Is Nothing) Then
         Set CreditStockCode = GetObject("CCreditStockCode", Cl, Trim(TempData.STOCK_NO), False)
         If CreditStockCode Is Nothing Then
            Set CreditStockCode = New CCreditStockCode
            CreditStockCode.STOCK_DESC = TempData.STOCK_DESC
            CreditStockCode.UNIT_NAME = TempData.UNIT_NAME
            CreditStockCode.UNIT_CHANGE_NAME = TempData.UNIT_CHANGE_NAME
            CreditStockCode.UNIT_AMOUNT = TempData.UNIT_AMOUNT

            Call Cl.add(CreditStockCode, Trim(TempData.STOCK_NO))
         End If
         CreditStockCode.STOCK_NO = TempData.STOCK_NO
         If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Or TempData.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.EXT_DISCOUNT_AMOUNT

            TotalSalePrice = TotalSalePrice + CreditStockCode.CREDIT_BALANCE
            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE + TempData.AVG_PRICE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT + TempData.TOTAL_AMOUNT

         ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE - TempData.TOTAL_PRICE

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.DISCOUNT_AMOUNT

            CreditStockCode.CREDIT_BALANCE = CreditStockCode.CREDIT_BALANCE + TempData.EXT_DISCOUNT_AMOUNT

            CreditStockCode.AVG_PRICE = CreditStockCode.AVG_PRICE - TempData.AVG_PRICE

            TotalSalePrice = TotalSalePrice - CreditStockCode.CREDIT_BALANCE

            CreditStockCode.TOTAL_AMOUNT = CreditStockCode.TOTAL_AMOUNT - TempData.TOTAL_AMOUNT

         End If
      End If

      Set TempData = Nothing
      Set CreditStockCode = Nothing
      Rs1.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   If Rs1.State = adStateOpen Then
      Rs1.Close
   End If
   Set Rs1 = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(29, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(29, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountAparStockCodeDocTypeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(36, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(36, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_CODE & "-" & TempData.STOCK_NO & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountAparGroupDocTypeStockCodeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(75, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(75, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_GROUP_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountAparTypeStockCodeFree(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String, Optional InCludeFree As Long = -1)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
      
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(78, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(78, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_TYPE_NAME & "-" & TempData.DOCUMENT_TYPE & "-" & TempData.STOCK_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
End Sub
Public Sub GetSaleAmountDocumentNo(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentTypeSet As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   BD.DOCUMENT_TYPE_SET = DocumentTypeSet
   
   Call BD.QueryData(1004, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1004, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.DOCUMENT_NO))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set BD = Nothing
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description & " " & TempData.DOCUMENT_NO & " ซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำซ้ำ"
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadMasterTypeName(C As ComboBox)
Dim I As Long
Dim j As Long
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   For I = 1 To 50
      If Len(MasterType2String(I)) > 0 Then
         j = j + 1
         C.AddItem (MasterType2String(I))
         C.ItemData(j) = I
      End If
   Next I
End Sub
Public Function MasterType2String(M As MASTER_TYPE) As String
   MasterType2String = ""
   If M = MASTER_COUNTRY Then
      MasterType2String = "ประเทศ"
   ElseIf M = MASTER_CUSTYPE Then
      MasterType2String = "ประเภทลูกค้า"
   ElseIf M = MASTER_SUPTYPE Then
      MasterType2String = "ประเภทซัพพลายเออร์"
   ElseIf M = MASTER_DEPARTMENT Then
      MasterType2String = "แผนก"
   ElseIf M = MASTER_UNIT Then
      MasterType2String = "หน่วย"
   ElseIf M = MASTER_STOCKTYPE Then
      MasterType2String = "ประเภทสต็อค"
   ElseIf M = MASTER_STOCKGROUP Then
      MasterType2String = "กลุ่มสต็อค"
   ElseIf M = MASTER_LOCATION Then
      MasterType2String = "คลัง"
   ElseIf M = MASTER_BANK Then
      MasterType2String = "ธนาคาร"
   ElseIf M = MASTER_BBRANCH Then
      MasterType2String = "สาขาธนาคาร"
   ElseIf M = MASTER_CHEQUE_TYPE Then
      MasterType2String = "ประเภทเช็ค"
   ElseIf M = MASTER_CNDN_REASON Then
      MasterType2String = "เหตุผลรับคืน"
   ElseIf M = MASTER_INVOICE_SUB Then
      MasterType2String = "ใบส่งสินค้าย่อย"
   ElseIf M = MASTER_INVOICE_RETURN Then
      MasterType2String = "ใบส่งสินค้าคืน"
   ElseIf M = MASTER_SUBTRACT Then
      MasterType2String = "ส่วนหัก"
   ElseIf M = MASTER_BANK_ACCOUNT Then
      MasterType2String = "เลขที่บัญชี"
   ElseIf M = MASTER_BACCOUNT_TYPE Then
      MasterType2String = "ประเภทบัญชี"
   ElseIf M = MASTER_ADDITION Then
      MasterType2String = "ส่วนเพิ่ม"
   ElseIf M = MASTER_CUSGROUP Then
      MasterType2String = "กลุ่มลูกค้า"
   ElseIf M = MASTER_INVENTORY_SUB_TYPE Then
      MasterType2String = "ประเภทเอกสารย่อย"
   End If
   
End Function
Public Sub LoadDisTinctBillingDocID(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DocumentNo As String, Optional DocumentType As Long, Optional FromDueDate As Date = -1, Optional ToDueDate As Date = -1, Optional OrderBy As Long = -1, Optional DocumentTypeSet As String, Optional BillingDocPack As Long)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.FROM_DUE_DATE = FromDueDate
   D.TO_DUE_DATE = ToDueDate
   D.DOCUMENT_NO = DocumentNo
   D.DOCUMENT_TYPE = DocumentType
   D.DOCUMENT_TYPE_SET = DocumentTypeSet
   D.ORDER_BY = OrderBy
   Call D.QueryData(1003, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(1003, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub LoadDisTinctPOID(Cl As Collection, Optional FromDate As Date = -1, Optional ToDate As Date = -1, Optional DocumentTypeSet As String)
On Error GoTo ErrorHandler
Dim D As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set D = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   D.BILLING_DOC_ID = -1
   D.FROM_DATE = FromDate
   D.TO_DATE = ToDate
   D.DOCUMENT_TYPE_SET = DocumentTypeSet
   Call D.QueryData(101, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(101, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.PO_ID)))
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
Public Sub GetSaleAmountEmpDocType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(102, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(102, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetSaleAmountEmpDocTypeBranch(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional InCludeFree As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_SALE_CODE = FromSaleCode
   BD.TO_SALE_CODE = ToSaleCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   
   Call BD.QueryData(103, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(103, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.SALE_CODE & "-" & TempData.DOCUMENT_TYPE))
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
Public Sub GetLotItemPartTxType(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Call Lt.QueryData(31, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(31, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE))
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
Public Sub GetLotItemPartTxType2(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim Lt As CLotItem
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLotItem
Dim m_Data As CMasterRef
Dim I As Long

   Set m_Data = New CMasterRef
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
         
         Set Lt = New CLotItem
         Set Rs = New ADODB.Recordset
         
         Lt.FROM_DOC_DATE = FromDate
         Lt.TO_DOC_DATE = ToDate
         Lt.FROM_STOCK_NO = FromStockNo
         Lt.TO_STOCK_NO = ToStockNo
         Call Lt.QueryData(41, Rs, ItemCount, False)
         I = 0
         While Not Rs.EOF
            I = I + 1
            Set TempData = New CLotItem
            Call TempData.PopulateFromRS(41, Rs)
            
            If Not (Cl Is Nothing) Then
               Call Cl.add(TempData, Trim(TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE & "-" & TempData.LOCATION_NO))
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
Public Sub GetSumAmountByPartItemIndSub(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   Call Lt.QueryData(34, Rs, ItemCount, False)
   
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
      Call TempData.PopulateFromRS(34, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INVENTORY_SUB_TYPE & "-" & TempData.PART_ITEM_ID))
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
Public Sub GetDistinctPartOutputByInput(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
      
   Call Lt.QueryData(44, Rs, ItemCount, False)
   
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
      
      Call TempData.PopulateFromRS(44, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetDistinctPartInputByInput(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1, Optional FromStockNo As String, Optional ToStockNo As String)
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
   Lt.DOCUMENT_TYPE = DocumentType
   Lt.FROM_STOCK_NO = FromStockNo
   Lt.TO_STOCK_NO = ToStockNo
   
      
   Call Lt.QueryData(46, Rs, ItemCount, False)
   
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
      
      Call TempData.PopulateFromRS(46, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData)
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
Public Sub GetSumAmountByInputOutput(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional LocationID As Long = -1, Optional DocumentType As Long = -1)
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
   Lt.DOCUMENT_TYPE = DocumentType
'   Lt.FROM_STOCK_NO = FromStockNo
'   Lt.TO_STOCK_NO = ToStockNo
   
   Call Lt.QueryData(47, Rs, ItemCount, False)
   
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
      
      Call TempData.PopulateFromRS(47, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.INVENTORY_DOC_ID & "-" & TempData.PART_ITEM_ID & "-" & TempData.TX_TYPE))
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
Public Sub GetSaleAmountCustomerStockGroup(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_APAR_CODE = FromCustomerCode
   BD.TO_APAR_CODE = ToCustomerCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(167, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(167, Rs)
      
      Set BD = GetObject("CBillingDoc", Cl, Trim(Str(TempData.APAR_MAS_ID)), False)
      If BD Is Nothing Then
         Set BD = New CBillingDoc
         BD.APAR_MAS_ID = TempData.APAR_MAS_ID
         BD.APAR_TYPE = TempData.APAR_TYPE
         Call Cl.add(BD, Trim(Str(TempData.APAR_MAS_ID)))
      End If
      
      Call BD.CollBillingDoc.add(TempData)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = "ข้อมูล Load : GetSaleAmountCustomerStockGroup มีข้อมูลซ้ำ "
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetSaleAmountByCustomer(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_APAR_CODE = FromCustomerCode
   BD.TO_APAR_CODE = ToCustomerCode
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(168, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(168, Rs)
      
      Call Cl.add(TempData, Trim(Str(TempData.APAR_MAS_ID)))
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = "ข้อมูล Load : GetSaleAmountByCustomer มีข้อมูลซ้ำ "
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadApArMasAddressLetter(Cl As Collection, Optional FromCustomerCode As String, Optional ToCustomerCode As String, Optional ApArType As Long, Optional ApArGroup As Long)
On Error GoTo ErrorHandler
Dim Apm As CAPARMas
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAPARMas
Dim I As Long
   
   Set Apm = New CAPARMas
   Set Rs = New ADODB.Recordset
   
   Apm.FROM_APAR_CODE = FromCustomerCode
   Apm.TO_APAR_CODE = ToCustomerCode
   Apm.APAR_TYPE = ApArType
   Apm.APAR_GROUP = ApArGroup
   Apm.APAR_IND = 1
   Call Apm.QueryData(5, Rs, ItemCount, False)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   
   
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CAPARMas
      Call TempData.PopulateFromRS(5, Rs)
      
      Call Cl.add(TempData)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = "ข้อมูล Load : LoadApArMasAddress มีข้อมูลซ้ำ " & " " & Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadFormula(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CFormula
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CFormula
Dim I As Long

   Set Rs = New ADODB.Recordset
   Set D = New CFormula
      
   D.FORMULA_ID = -1
   Call D.QueryData(1, Rs, ItemCount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CFormula
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.FORMULA_DESC)
         C.ItemData(I) = TempData.FORMULA_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.FORMULA_ID)))
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
Public Sub GetAllDetailJobItem(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromStockNo As String, Optional ToStockNo As String)
On Error GoTo ErrorHandler
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.FROM_STOCK_NO = FromStockNo
   BD.TO_STOCK_NO = ToStockNo
   BD.DOCUMENT_TYPE_SET = "(" & RECEIPT1_DOCTYPE & "," & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   Call BD.QueryData(167, Rs, ItemCount)
   
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
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS2(167, Rs)
      
      Set BD = GetObject("CBillingDoc", Cl, Trim(Str(TempData.APAR_MAS_ID)), False)
      If BD Is Nothing Then
         Set BD = New CBillingDoc
         BD.APAR_MAS_ID = TempData.APAR_MAS_ID
         BD.APAR_TYPE = TempData.APAR_TYPE
         Call Cl.add(BD, Trim(Str(TempData.APAR_MAS_ID)))
      End If
      
      Call BD.CollBillingDoc.add(TempData)
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = "ข้อมูล Load : GetSaleAmountCustomerStockGroup มีข้อมูลซ้ำ "
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

