VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing
   
   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub
'==
Public Function QueryEnterprise(Ua As CEnterprise, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryEnterprise"
   QueryEnterprise = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CEnterpriseAddress
      Set cEnpAddr = New CEnterpriseAddress
      Call cEnpAddr.SetFieldValue("ENTERPRISE_ID", Ua.ENTERPRISE_ID)
      Call cEnpAddr.SetFieldValue("ORDER_BY", 1)
      Call cEnpAddr.QueryData(1, m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.EnpAddresses = Nothing
      Set Ua.EnpAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CEnterpriseAddress
         Call cEnpAddr.PopulateFromRS(1, m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         Call cAddr.SetFieldValue("ADDRESS_ID", cEnpAddr.GetFieldValue("ADDRESS_ID"))
         Call cAddr.QueryData(1, m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(1, m_Rs2)
            cAddr.Flag = "I"
            Call cEnpAddr.Addresses.add(cAddr)
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.EnpAddresses.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If

   QueryEnterprise = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryEnterprise = False
End Function

Public Function AddEditEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim Name As CName
Dim Addr1 As CAddress

   RName = "AddEditEnterprise"
   AddEditEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
         
   Dim EmpAddr As CEnterpriseAddress
   Dim Addr As CAddress
   
   For Each EmpAddr In Ug.EnpAddresses
      If EmpAddr.Flag = "A" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.ShowMode = SHOW_ADD
         Call Addr.AddEditData
                  
         Call EmpAddr.SetFieldValue("ENTERPRISE_ID", Ug.ENTERPRISE_ID)
         Call EmpAddr.SetFieldValue("ADDRESS_ID", Addr.GetFieldValue("ADDRESS_ID"))
         EmpAddr.ShowMode = SHOW_ADD
         EmpAddr.AddEditData
      ElseIf EmpAddr.Flag = "E" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.ShowMode = SHOW_EDIT
         Call Addr.AddEditData
         
      ElseIf EmpAddr.Flag = "D" Then
         Call EmpAddr.DeleteData
         
         Set Addr = EmpAddr.Addresses(1)
         Call Addr.DeleteData
      End If
   Next EmpAddr
   
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
      
   HasBegin = False
   IsOK = True

   AddEditEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEnterprise = False
End Function

Public Function DeleteEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteEnterprise"
   DeleteEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_MSGBOX)
   
   IsOK = False

   DeleteEnterprise = True
End Function

Public Function QueryCustomer(Ua As CAPARMas, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCustomer"
   QueryCustomer = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CApArAddress
      Set cEnpAddr = New CApArAddress
      Call cEnpAddr.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
      Call cEnpAddr.SetFieldValue("ORDER_BY", 1)
      Call cEnpAddr.QueryData(1, m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddresses = Nothing
      Set Ua.CstAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CApArAddress
         Call cEnpAddr.PopulateFromRS(1, m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         Call cAddr.SetFieldValue("ADDRESS_ID", cEnpAddr.GetFieldValue("ADDRESS_ID"))
         Call cAddr.QueryData(1, m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(1, m_Rs2)
'            Call cAddr.SetFieldValue("ADDRESS_TYPE", cEnpAddr.GetFieldValue("ADDRESS_TYPE"))
            cAddr.Flag = "I"
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
         
         cEnpAddr.Flag = "I"
         Call Ua.CstAddresses.add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If
      
   QueryCustomer = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomer = False
End Function

Public Function AddEditCustomer(Ug As CAPARMas, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomer"
   AddEditCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
      
   Dim CstAddr As CApArAddress
   Dim Addr As CAddress
   
   For Each CstAddr In Ug.CstAddresses
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.ShowMode = SHOW_ADD
         Call Addr.AddEditData

         Call CstAddr.SetFieldValue("APAR_MAS_ID", Ug.APAR_MAS_ID)
         Call CstAddr.SetFieldValue("ADDRESS_ID", Addr.GetFieldValue("ADDRESS_ID"))
'         Call CstAddr.SetFieldValue("ADDRESS_TYPE", -1)
         CstAddr.ShowMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.ShowMode = SHOW_EDIT
         Call Addr.AddEditData

         Call CstAddr.SetFieldValue("APAR_MAS_ID", Ug.APAR_MAS_ID)
         Call CstAddr.SetFieldValue("ADDRESS_ID", Addr.GetFieldValue("ADDRESS_ID"))
'         Call CstAddr.SetFieldValue("ADDRESS_TYPE", Addr.GetFieldValue("ADDRESS_TYPE"))
         CstAddr.ShowMode = SHOW_EDIT
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   Dim CUS As CAPARMas
   Set CUS = New CAPARMas
   CUS.APAR_IND = 1
   Call LoadApArMas(CUS, Nothing, m_CustomerColl)
   Set CUS = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomer = False
End Function

Public Function DeleteCustomer(Ug As CAPARMas, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCustomer"
   DeleteCustomer = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteCustomer = False
End Function

Public Function DeleteMasterRef(Ug As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteMasterRef"
   DeleteMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteMasterRef = True
End Function

Public Function QueryEmployee(Ua As CEmployee, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

Dim Qp As CEmployee

   RName = "QueryEmployee"
   QueryEmployee = False
   
   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      
   End If
   
   QueryEmployee = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmployee = False
End Function

Public Function DeleteEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteEmployee"
   DeleteEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteEmployee = True
End Function

Public Function AddEditEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEmployee"
   AddEditEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim TempEmp As CEmployee
   Set TempEmp = New CEmployee
   Call LoadEmployee(TempEmp, Nothing, m_EmployeeColl)
   Set TempEmp = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEmployee = False
End Function

Public Function QueryUserAccount(Ua As CUserAccount, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryUserAccount"
   QueryUserAccount = False

   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryUserAccount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserAccount = False
End Function

Public Function QueryInventoryDoc(Ua As CInventoryDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryInventoryDoc"
   QueryInventoryDoc = False
   
   IsOK = True
   
   Call Ua.QueryData(1, Rs, ItemCount, True)
   If Ua.QueryFlag = 1 Then
      Dim Gr As CLotItem
      Set Gr = New CLotItem
      Gr.CANCEL_FLAG = ""
      Gr.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
      Call Gr.QueryData(1, m_Rs1, iCount, False)
      Set Gr = Nothing
      
      Set Ua.ImportExportItems = Nothing
      Set Ua.ImportExportItems = New Collection
      
      While Not m_Rs1.EOF
         Set Gr = New CLotItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Call Ua.ImportExportItems.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If

   QueryInventoryDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryDoc = False
End Function

Public Function QueryStockCode(Ua As CStockCode, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
   RName = "QueryStockCode"
   QueryStockCode = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount, True)
   
   If Ua.QueryFlag = 1 Then
      
   End If
   
   QueryStockCode = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryStockCode = False
End Function

Public Function AddEditUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserAccount"
   AddEditUserAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserAccount = False
End Function


Public Function AddEditInventoryDoc(Ua As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditInventoryDoc"
   AddEditInventoryDoc = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Cd As CConfigDoc
   If Ua.ShowMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ua.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ua.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ua.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim Gr As CLotItem
   For Each Gr In Ua.ImportExportItems
      Gr.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
      Gr.CANCEL_FLAG = Ua.CANCEL_FLAG
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If (Ua.OLD_CANCEL_FLAG <> Ua.CANCEL_FLAG) And Len(Ua.OLD_CANCEL_FLAG) > 0 Then
      For Each Gr In Ua.ImportExportItems
         Gr.ShowMode = SHOW_EDIT
         Gr.CANCEL_FLAG = Ua.CANCEL_FLAG
         Call Gr.AddEditData
      Next Gr
   End If
   Set Gr = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditInventoryDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditInventoryDoc = False
End Function
Public Function AddEditFormula(Ua As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditFormula"
   AddEditFormula = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim FmlIp As CFormulaItem
   For Each FmlIp In Ua.CollFormulaInputs
      FmlIp.FORMULA_ID = Ua.FORMULA_ID
      FmlIp.TX_TYPE = "I"
      If FmlIp.Flag = "A" Then
         FmlIp.AddEditMode = SHOW_ADD
         Call FmlIp.AddEditData
      ElseIf FmlIp.Flag = "E" Then
         FmlIp.AddEditMode = SHOW_EDIT
         Call FmlIp.AddEditData
      ElseIf FmlIp.Flag = "D" Then
         Call FmlIp.DeleteData
      End If
   Next FmlIp
   
   For Each FmlIp In Ua.CollFormulaOutputs
      FmlIp.FORMULA_ID = Ua.FORMULA_ID
      FmlIp.TX_TYPE = "O"
      If FmlIp.Flag = "A" Then
         FmlIp.AddEditMode = SHOW_ADD
         Call FmlIp.AddEditData
      ElseIf FmlIp.Flag = "E" Then
         FmlIp.AddEditMode = SHOW_EDIT
         Call FmlIp.AddEditData
      ElseIf FmlIp.Flag = "D" Then
         Call FmlIp.DeleteData
      End If
   Next FmlIp
   
   Set FmlIp = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   AddEditFormula = False
End Function
Public Function QueryFormula(Ua As CFormula, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryFormula"
   QueryFormula = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim FmlI As CFormulaItem
      Set FmlI = New CFormulaItem
      FmlI.FORMULA_ID = Ua.FORMULA_ID
      Call FmlI.QueryData(1, m_Rs1, iCount)
      Set FmlI = Nothing

      Set Ua.CollFormulaInputs = Nothing
      Set Ua.CollFormulaInputs = New Collection
      
      While Not m_Rs1.EOF
         Set FmlI = New CFormulaItem
         Call FmlI.PopulateFromRS(1, m_Rs1)
         
         FmlI.Flag = "I"
         If FmlI.TX_TYPE = "I" Then             'เป็นวัตถุดิบ แต่เวลาไปเป็น STOCK จะต้อง "E" เนื่องจากเบิกออกจาก STOCK
            Call Ua.CollFormulaInputs.add(FmlI)
         Else
            Call Ua.CollFormulaOutputs.add(FmlI)
         End If
         Set FmlI = Nothing
         
         m_Rs1.MoveNext
      Wend
      
   End If

   QueryFormula = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryFormula = False
End Function

Public Function DeleteFormula(Ua As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteFormula"
   DeleteFormula = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteFormula = True
End Function

Public Function CopyFormula(Ug As CFormula, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyFormula"
   CopyFormula = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryFormula(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.FORMULA_ID = -1
   Ug.FORMULA_NO = "***" & Ug.FORMULA_NO
   
   Dim Det As CFormulaItem
   For Each Det In Ug.CollFormulaInputs
      Det.Flag = "A"
   Next Det
      
   For Each Det In Ug.CollFormulaOutputs
      Det.Flag = "A"
   Next Det
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditFormula(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyFormula = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyFormula = False
End Function

Public Function AddEditStockCode(Ua As CStockCode, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditStockCode"
   AddEditStockCode = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditStockCode = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditStockCode = False
End Function

Public Function AddEditMasterRef(Ua As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterRef"
   AddEditMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterRef = False
End Function

Public Function DeleteUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteUserAccount"
   DeleteUserAccount = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteUserAccount = True
End Function

Public Function DeleteInventoryDoc(Ua As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteInventoryDoc"
   DeleteInventoryDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteInventoryDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteInventoryDoc = True
End Function

Public Function DeleteStockCode(Ua As CStockCode, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteStockCode"
   DeleteStockCode = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteStockCode = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteStockCode = True
End Function

Public Function DBLogin(UserName As String, Password As String, CanLogin As Boolean, User As CUser, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim ItemCount As Long
Dim LoginID As Long

   RName = "Login"
   DBLogin = False

   SQL1 = "SELECT COUNT(*) FROM USER_ACCOUNT "
   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   If NVLI(m_Rs1(0), 0) <= 0 Then 'ถ้าไม่มีผู้ใช้อยุในดาตาเบสให้ล็อคอินอะไรก็ได้
      If UserName <> "" Then
         ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

         DBLogin = True
         CanLogin = False
         Exit Function
      End If
      
      User.USER_ID = 0
      User.GROUP_ID = 0
      User.GROUP_NAME = ""
      User.USER_NAME = UserName

      DBLogin = True
      CanLogin = True
      Exit Function
   End If

   SQL1 = "SELECT UD.*, GD.* FROM USER_ACCOUNT UD, USER_GROUP GD "
   SQL1 = SQL1 & "WHERE (GD.GROUP_ID = UD.GROUP_ID) "
   SQL1 = SQL1 & "AND (UD.USER_NAME = '" & ChangeQuote(UserName) & "' ) "
   SQL1 = SQL1 & "AND (UD.USER_PASSWORD = '" & ChangeQuote(EncryptText(Password)) & "' ) "

   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

      DBLogin = True
      CanLogin = False
      Exit Function
   End If

   If NVLS(m_Rs1("USER_STATUS"), "Y") <> "Y" Then
      ErrorObj.LocalErrorMsg = "รายชื่อของท่านขณะนี้ไม่สามารถเข้าใช้งานระบบได้ โปรดติดต่อผู้บริหารระบบ "

      DBLogin = True
      CanLogin = False
      Exit Function
   End If
   
   User.GROUP_ID = NVLI(m_Rs1("GROUP_ID"), -1)
   User.GROUP_NAME = NVLS(m_Rs1("GROUP_NAME"), "")
   User.USER_NAME = UserName
   User.REAL_USER_ID = NVLI(m_Rs1("USER_ID"), -1)
   User.USER_ID = NVLI(m_Rs1("USER_ID"), -1)
   
'   m_ADOConn.BeginTrans
'
'   glbLoginTracking.LOGIN_DATE = Now
'   glbLoginTracking.LOGOUT_DATE = -1
'   glbLoginTracking.IP_ADDRESS = glbDatabaseMngr.m_Winsock.LocalIP
'   glbLoginTracking.USER_NAME = User.USER_NAME
'   glbLoginTracking.USER_ID = User.USER_ID
'   glbLoginTracking.AddEditMode = SHOW_ADD
'   Call glbLoginTracking.AddEditData
'
'   m_ADOConn.CommitTrans
   
   CanLogin = True
   DBLogin = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   DBLogin = False
End Function
Public Function QueryUserGroup(Ua As CUserGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean

   RName = "QueryUserGroup"
   QueryUserGroup = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CGroupRight
      Set Gr = New CGroupRight
      Call Gr.SetFieldValue("GROUP_ID", Ua.GetFieldValue("GROUP_ID"))
      Call Gr.QueryData(2, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.RightItems = Nothing
      Set Ua.RightItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CGroupRight
         Call Gr.PopulateFromRS(2, m_Rs1)
         Call Ua.RightItems.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryUserGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserGroup = False
End Function

Public Function AddEditUserGroup(Ua As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserGroup"
   AddEditUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CGroupRight
   For Each Gr In Ua.RightItems
      Call Gr.SetFieldValue("GROUP_ID", Ua.GetFieldValue("GROUP_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserGroup = False
End Function

Public Function DeleteUserGroup(Ug As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteUserGroup"
   DeleteUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   
   IsOK = False

   DeleteUserGroup = True
End Function

Public Function ChangePassword(UserID As Long, OldPassword As String, NewPassword As String, IsCanChange As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim UserName As String
Dim ItemCount As Long

   RName = "ChangePassword"
   ChangePassword = False
   HasBegin = False

   SQL = "SELECT * FROM USER_ACCOUNT WHERE (USER_ID = " & UserID & ") "
   If Not glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, m_Rs1, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่มีผู้ใช้นี้อยู่ในฐานข้อมูล "
      IsCanChange = False
      ChangePassword = True
      Exit Function
   End If
   UserName = NVLS(m_Rs1("USER_NAME"), "")

   SQL = "SELECT * FROM USER_ACCOUNT "
   SQL = SQL & "WHERE (USER_PASSWORD = '" & ChangeQuote(EncryptText(OldPassword)) & "') AND "
   SQL = SQL & "(USER_ID = " & UserID & ") "
   If Not glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, m_Rs1, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   If m_Rs1.EOF Then
      ErrorObj.LocalErrorMsg = "ท่านใส่รหัสผ่านเก่าไม่ถูกต้อง "
      IsCanChange = False
      ChangePassword = True
      Exit Function
   End If

   m_ADOConn.BeginTrans
   HasBegin = True

   SQL = "UPDATE USER_ACCOUNT SET "
   SQL = SQL & "USER_PASSWORD = '" & ChangeQuote(EncryptText(NewPassword)) & "' "
   SQL = SQL & "WHERE (USER_ID = " & UserID & ") "
   Call m_ADOConn.Execute(SQL)

   Call m_ADOConn.CommitTrans
   HasBegin = False

   IsCanChange = True
   ChangePassword = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description & SQL
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = ""
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   ChangePassword = False
End Function

Public Sub CreateTransferItems(Ivd As CInventoryDoc)
Dim Ti As CTransferItem
Dim O As Object
Dim Ei As CLotItem
Dim II As CLotItem
Dim I As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.ImportExportItems.Count \ 2
   Count2 = Ivd.ImportExportItems.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = Ivd.ImportExportItems(j)
         If (O.TX_TYPE = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.TX_TYPE = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set Ei = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = Ei
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.add(Ti)
      Set Ti = Nothing
   Next I
End Sub
Public Function QueryBillingDoc(Ua As CBillingDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBillingDoc"
   QueryBillingDoc = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   If Ua.QueryFlag = 1 Then
      Dim Gr As CDocItem
      Set Gr = New CDocItem
      
      Dim iCount2 As Long
      
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DocItems = Nothing
      Set Ua.DocItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CDocItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.DocItems.add(Gr)
         Set Gr = Nothing
         
         m_Rs1.MoveNext
      Wend
      Set Gr = Nothing
      
      Dim Rcp As CRcpCnDn_Item
      Set Rcp = New CRcpCnDn_Item
      Rcp.CheckBillsFlag = Ua.CHECK_BILLS_FLAG
      Call Rcp.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Rcp.QueryData(1, m_Rs1, iCount)
      Set Rcp = Nothing
      
      Set Ua.RcpCnDnItems = Nothing
      Set Ua.RcpCnDnItems = New Collection

      While Not m_Rs1.EOF
         Set Rcp = New CRcpCnDn_Item
         Call Rcp.PopulateFromRS(1, m_Rs1)
         Rcp.Flag = "I"
         'Debug.Print (Trim(Str(Rcp.GetFieldValue("DOC_ID"))))
         Call Ua.RcpCnDnItems.add(Rcp, Trim(Str(Rcp.GetFieldValue("DOC_ID"))))
         
         Set Rcp = Nothing

         m_Rs1.MoveNext
      Wend
      Set Rcp = Nothing
   
      Dim BSub As CBillingSubTract
      Set BSub = New CBillingSubTract
      Call BSub.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BSub.QueryData(1, m_Rs1, iCount)
      Set BSub = Nothing
      
      Set Ua.BillingSubTracts = Nothing
      Set Ua.BillingSubTracts = New Collection

      While Not m_Rs1.EOF
         Set BSub = New CBillingSubTract
         Call BSub.PopulateFromRS(1, m_Rs1)
         BSub.Flag = "I"
         Call Ua.BillingSubTracts.add(BSub)
         Set BSub = Nothing

         m_Rs1.MoveNext
      Wend
      Set BSub = Nothing
      
      Dim BAdd As CBillingAddition
      Set BAdd = New CBillingAddition
      Call BAdd.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BAdd.QueryData(1, m_Rs1, iCount)
      Set BAdd = Nothing
      
      Set Ua.BillingAdditions = Nothing
      Set Ua.BillingAdditions = New Collection
      
      While Not m_Rs1.EOF
         Set BAdd = New CBillingAddition
         Call BAdd.PopulateFromRS(1, m_Rs1)
         BAdd.Flag = "I"
         Call Ua.BillingAdditions.add(BAdd)
         Set BAdd = Nothing

         m_Rs1.MoveNext
      Wend
      Set BAdd = Nothing
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.Payments = Nothing
      Set Ua.Payments = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.Payments.add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryBillingDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBillingDoc = False
End Function
Public Function QueryBillingDocOnlyDo(Ua As CBillingDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim iCount As Long

   RName = "QueryBillingDocOnlyDo"
   QueryBillingDocOnlyDo = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   If Ua.QueryFlag = 1 Then
      Dim Gr As CDocItem
      Set Gr = New CDocItem
      
      Dim iCount2 As Long
      
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DocItems = Nothing
      Set Ua.DocItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CDocItem
         Call Gr.PopulateFromRS(1, m_Rs1)
                  
         Gr.Flag = "I"
         Call Ua.DocItems.add(Gr)
         Set Gr = Nothing
         
         m_Rs1.MoveNext
      Wend
      Set Gr = Nothing
      
   End If
   
   
   
   QueryBillingDocOnlyDo = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryBillingDocOnlyDo = False
End Function

Public Function QueryCheque(Ua As CCheque, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCheque"
   QueryCheque = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryCheque = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCheque = False
End Function

Public Function DeleteCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCheque"
   DeleteCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCheque = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
'   ErrorObj.SystemErrorMsg = Err.Description
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   
   IsOK = False

   DeleteCheque = True
End Function

Public Function DeleteBillingDoc(Ua As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim BD As CBillingDoc
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long

   RName = "DeleteBillingDoc"
   DeleteBillingDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Set BD = New CBillingDoc
   Set m_Rs = New ADODB.Recordset
   
   BD.BILLING_DOC_ID = Ua.BILLING_DOC_ID
   BD.QueryFlag = 1
   Call QueryBillingDoc(BD, m_Rs, ItemCount, IsOK, glbErrorLog)
   
   Call BD.PopulateFromRS(1, m_Rs)
   
   Ua.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
   Call Ua.DeleteData
   
   'Call UpDateBDRcpCnDnItem(BD) 'ปิดไว้เนื่องจากว่ามนไม่จำเป็นที่จะต้องดูที่หน้านี้ดูจากรายงานน่าจะดีกว่า
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBillingDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteBillingDoc = True
End Function

Public Function AddEditBillingDoc(Ua As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PaymentType As Long
Dim OldPaymentType As Long
Dim I As Long
   RName = "AddEditBillingDoc"
   AddEditBillingDoc = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Cd As CConfigDoc
    If Ua.ShowMode = SHOW_ADD And Ua.RUNNING_NO > 0 Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ua.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ua.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ua.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim Gr As CDocItem
   For Each Gr In Ua.DocItems
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   Set Gr = Nothing
   
   Dim Rc As CRcpCnDn_Item
   For Each Rc In Ua.RcpCnDnItems
      Call Rc.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Rc.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      If Rc.Flag = "A" Then
         Rc.ShowMode = SHOW_ADD
         Call Rc.AddEditData
      ElseIf Rc.Flag = "E" Then
         Rc.ShowMode = SHOW_EDIT
         Call Rc.AddEditData
      ElseIf Rc.Flag = "D" Then
         Call Rc.DeleteData
      End If
   Next Rc
   Set Rc = Nothing
   
   Dim BSub As CBillingSubTract
   For Each BSub In Ua.BillingSubTracts
      Call BSub.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BSub.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      If BSub.Flag = "A" Then
         BSub.ShowMode = SHOW_ADD
         Call BSub.AddEditData
      ElseIf BSub.Flag = "E" Then
         BSub.ShowMode = SHOW_EDIT
         Call BSub.AddEditData
      ElseIf BSub.Flag = "D" Then
         Call BSub.DeleteData
      End If
   Next BSub
   Set BSub = Nothing
   
   Dim BAdd As CBillingAddition
   For Each BAdd In Ua.BillingAdditions
      Call BAdd.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call BAdd.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      If BAdd.Flag = "A" Then
         BAdd.ShowMode = SHOW_ADD
         Call BAdd.AddEditData
      ElseIf BAdd.Flag = "E" Then
         BAdd.ShowMode = SHOW_EDIT
         Call BAdd.AddEditData
      ElseIf BAdd.Flag = "D" Then
         Call BAdd.DeleteData
      End If
   Next BAdd
   Set BAdd = Nothing
   
   
   Dim Ct As CCashTran
   Dim TempChequeID As Long
   For Each Ct In Ua.Payments
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ct.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
      Call Ct.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
      Call Ct.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
      Call Ct.SetFieldValue("EMP_ID", Ua.SALE_BY)
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Or _
         Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And Ct.GetFieldValue("TX_TYPE") = "I" Then 'เช็ค
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_TYPE", Ua.DOCUMENT_TYPE)
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
            
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
         End If
         If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And Ct.GetFieldValue("TX_TYPE") = "E" Then 'เช็คออกจากมือ
            Call Ct.SetFieldValue("CHEQUE_ID", TempChequeID)
         Else
            Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            TempChequeID = Ct.GetFieldValue("CHEQUE_ID")
         End If
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And Not (OldPaymentType = CHEQUE_HAND_PMT Or OldPaymentType = CHEQUE_BANK_PMT) _
         And (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT) And Ct.GetFieldValue("TX_TYPE") = "I" Then
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And Ct.GetFieldValue("TX_TYPE") = "E" Then 'เช็คออกจากมือ
               Call Ct.SetFieldValue("CHEQUE_ID", TempChequeID)
            Else
               Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
               TempChequeID = Ct.GetFieldValue("CHEQUE_ID")
            End If
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = CHEQUE_HAND_PMT Or OldPaymentType = CHEQUE_BANK_PMT) _
         And Not (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT) Then
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHEQUE_ID", -1)
            Call Ct.AddEditData
            
            Call Ct.Cheque.DeleteData
         ElseIf ((PaymentType = OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT)) _
         Or ((PaymentType <> OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT) _
         And (OldPaymentType = CHEQUE_HAND_PMT Or OldPaymentType = CHEQUE_BANK_PMT)) Then
            Call Ct.Cheque.SetFieldValue("TX_NO", Ua.DOCUMENT_NO)
            Call Ct.Cheque.SetFieldValue("TX_DATE", Ua.DOCUMENT_DATE)
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.SetFieldValue("BILLING_DOC_PACK", Ua.BILLING_DOC_PACK)
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If Ct.GetFieldValue("OLD_CHEQUE_ID") > 0 Then 'เช็ค
            Call Ct.Cheque.DeleteData
         End If
      End If
   Next Ct

   'ทำเพื่อเซต APAR_ID ให้ตรงกันกับที่ BILLING_DOC
   For Each Ct In Ua.Payments
      If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
         Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
         Call Ct.Cheque.SetFieldValue("APAR_MAS_ID", Ua.APAR_MAS_ID)
         Call Ct.Cheque.UpdateApArID
      End If
   Next Ct
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditBillingDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBillingDoc = False
End Function

Public Function AddEditCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCheque"
   AddEditCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCheque = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCheque = False
End Function

Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(1, Rs, ItemCount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData(Ug.ShowMode)

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReportConfig = False
End Function
Public Function DO2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Area As Long, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDocItem
Dim Ei As CLotItem
Dim PartItemID As Long
   
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   DO2InventoryDoc = True
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
            
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
   
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.DOCUMENT_TYPE = DocumentType
   Ivd.CANCEL_FLAG = BD.CANCEL_FLAG
   Ivd.APAR_MAS_ID = BD.APAR_MAS_ID
   For Each Di In BD.DocItems
      PartItemID = Di.GetFieldValue("PART_ITEM_ID")
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         If Area = 1 Then
            Ei.TX_TYPE = "E"
            Ei.MULTIPLIER = -1
            
         ElseIf Area = 2 Then
            Ei.TX_TYPE = "I"
            Ei.MULTIPLIER = 1
            Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
            Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
         End If
         
         Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
         Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
                     
         Ei.Flag = "A"
         
         Ei.PART_ITEM_ID = PartItemID
         
         Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
         Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
         
         Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
         Ei.CANCEL_FLAG = BD.CANCEL_FLAG
         
         If PartItemID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExportItems.add(Ei)
         End If
         Set Ei = Nothing
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
         If Ei Is Nothing Then
            'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If PartItemID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItem
               If Area = 1 Then
                  Ei.TX_TYPE = "E"
                  Ei.MULTIPLIER = -1
               ElseIf Area = 2 Then
                  Ei.TX_TYPE = "I"
                  Ei.MULTIPLIER = 1
                  Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
                  Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
               End If
               Ei.Flag = "E"
               
               Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
               Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
         
               Ei.PART_ITEM_ID = PartItemID
               Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
               Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
               
               Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
               Ei.CANCEL_FLAG = BD.CANCEL_FLAG
               
               Call Ivd.ImportExportItems.add(Ei)
               Set Ei = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If PartItemID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Di.Flag
               If Area = 1 Then
                  Ei.TX_TYPE = "E"
                  Ei.MULTIPLIER = -1
               ElseIf Area = 2 Then
                  Ei.TX_TYPE = "I"
                  Ei.MULTIPLIER = 1
                  Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
                  Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
               End If
               
               Ei.PART_ITEM_ID = PartItemID
               
               Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
               Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
               
               Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
               
               Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
               
               Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
               Ei.CANCEL_FLAG = BD.CANCEL_FLAG
               
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function Rt2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Area As Long, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDocItem
Dim Ei As CLotItem
Dim PartItemID As Long
   
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   Rt2InventoryDoc = True
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
            
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
   
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.DOCUMENT_TYPE = DocumentType
   Ivd.CANCEL_FLAG = BD.CANCEL_FLAG
   Ivd.APAR_MAS_ID = BD.APAR_MAS_ID
   
   For Each Di In BD.DocItems
      PartItemID = Di.GetFieldValue("PART_ITEM_RETURN_ID")
      If Di.Flag = "A" Then
         Set Ei = New CLotItem
         
         If Area = 2 Then     ' ส่งคืนสินค้า
            Ei.TX_TYPE = "E"
            Ei.MULTIPLIER = -1
         ElseIf Area = 1 Then   ' รับคืนสินค้า
            Ei.TX_TYPE = "I"
            Ei.MULTIPLIER = 1
            
            Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
            Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
            
         End If
         
         Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
         Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
         
         Ei.Flag = "A"
         
         Ei.PART_ITEM_ID = PartItemID
         
         Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
         
         Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
         
         Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
         Ei.CANCEL_FLAG = BD.CANCEL_FLAG
         
         If PartItemID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExportItems.add(Ei)
         End If
         Set Ei = Nothing
      ElseIf Di.Flag = "E" Then
         Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
         If Ei Is Nothing Then
            'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If PartItemID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set Ei = New CLotItem
               If Area = 2 Then
                  Ei.TX_TYPE = "E"
                  Ei.MULTIPLIER = -1
                  
               ElseIf Area = 1 Then
                  Ei.TX_TYPE = "I"
                  Ei.MULTIPLIER = 1
                  Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
                  Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
               End If
               Ei.Flag = "E"
               
               Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
               Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
         
               Ei.PART_ITEM_ID = PartItemID
               Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
               
               Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
               
               Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
               Ei.CANCEL_FLAG = BD.CANCEL_FLAG
               
               Call Ivd.ImportExportItems.add(Ei)
               Set Ei = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If PartItemID > 0 Then
               'แก้ไขปกติ
               Ei.Flag = Di.Flag
               If Area = 2 Then
                  Ei.TX_TYPE = "E"
                  Ei.MULTIPLIER = -1
                  
               ElseIf Area = 1 Then
                  Ei.TX_TYPE = "I"
                  Ei.MULTIPLIER = 1
                  
                  Ei.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")
                  Ei.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
                  
               End If
               
               Ei.PART_ITEM_ID = PartItemID
               
               Ei.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
               Ei.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
               
               Ei.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
               
               If Di.GetFieldValue("LOT_ITEM_FLAG") = "N" Then
                  Ei.TX_AMOUNT = 0
               Else
                  Ei.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
               End If
               
               Ei.LINK_ID = Di.GetFieldValue("LINK_ID")
               Ei.CANCEL_FLAG = BD.CANCEL_FLAG
               
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               Ei.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"))
         If Not (Ei Is Nothing) Then
            Ei.Flag = "D"
         End If
      End If
   Next Di
End Function

Public Function RO2InventoryDoc(BD As CBillingDoc, Ivd As CInventoryDoc, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDocItem
   
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   If BD.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = BD.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
         
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
      
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = BD.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = BD.DOCUMENT_NO
   Ivd.DOCUMENT_TYPE = DocumentType
   Ivd.CANCEL_FLAG = BD.CANCEL_FLAG
   Ivd.APAR_MAS_ID = BD.APAR_MAS_ID
   
   Dim II As CLotItem
   
   For Each Di In BD.DocItems
      If Di.Flag = "A" Then
         Set II = New CLotItem
   
         II.Flag = "A"
         
         II.TX_TYPE = "I"
         II.MULTIPLIER = 1
         
         II.TOTAL_INCLUDE_PRICE = (Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")) - Di.GetFieldValue("DISCOUNT_AMOUNT") - Di.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
         II.AVG_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, Di.GetFieldValue("ITEM_AMOUNT"))
         
         II.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
         II.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
         
         II.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
         II.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
         II.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
         II.LINK_ID = Di.GetFieldValue("LINK_ID")
         II.CANCEL_FLAG = BD.CANCEL_FLAG
         
         
         If Di.GetFieldValue("PART_ITEM_ID") > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExportItems.add(II)
         End If
      ElseIf Di.Flag = "E" Then
         Set II = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"), "I")
         If II Is Nothing Then
            'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If Di.GetFieldValue("PART_ITEM_ID") > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               
               II.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
               II.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
               
               II.TX_TYPE = "I"
               II.MULTIPLIER = 1
               
               II.TOTAL_INCLUDE_PRICE = (Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")) - Di.GetFieldValue("DISCOUNT_AMOUNT") - Di.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
               II.AVG_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, Di.GetFieldValue("ITEM_AMOUNT"))
               
               II.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
               II.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
               II.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
               II.LINK_ID = Di.GetFieldValue("LINK_ID")
               II.CANCEL_FLAG = BD.CANCEL_FLAG
               
               
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If Di.GetFieldValue("PART_ITEM_ID") > 0 Then
               'แก้ไขปกติ
               II.Flag = Di.Flag
                                 
               II.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
               II.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
               
               II.TX_TYPE = "I"
               II.MULTIPLIER = 1
               
               II.TOTAL_INCLUDE_PRICE = (Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("ITEM_AMOUNT")) - Di.GetFieldValue("DISCOUNT_AMOUNT") - Di.GetFieldValue("EXTRA_DISCOUNT_AMOUNT")
               II.AVG_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, Di.GetFieldValue("ITEM_AMOUNT"))
               
               II.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
               II.LOCATION_ID = Di.GetFieldValue("LOCATION_ID")
               II.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
               'II.LINK_ID = Di.GetFieldValue("LINK_ID"))  ควรจะเป็นเลขเดิม
               II.CANCEL_FLAG = BD.CANCEL_FLAG
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               II.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set II = GetExportItem(Ivd, Di.GetFieldValue("LINK_ID"), "I")
         If Not (II Is Nothing) Then
            II.Flag = "D"
         End If
      End If
   Next Di
End Function
Public Function AddEditPackage(Ua As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditPackage"
   AddEditPackage = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CPackageDetail
   For Each Gr In Ua.PackageDeTails
      Call Gr.SetFieldValue("PACKAGE_ID", Ua.GetFieldValue("PACKAGE_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr

   Dim Package As CPackageDetail
   Set Package = New CPackageDetail
   Call Package.SetFieldValue("PACKAGE_DETAIL_ID", 1)
   Call LoadPackageDetail(Package, Nothing, LoadPackageColl)
      
   Set Package = Nothing
   
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditPackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPackage = False
End Function

Public Function QueryPackage(Ua As CPackage, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPackage"
   QueryPackage = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CPackageDetail
      Set Gr = New CPackageDetail
      Call Gr.SetFieldValue("PACKAGE_ID", Ua.GetFieldValue("PACKAGE_ID"))
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.PackageDeTails = Nothing
      Set Ua.PackageDeTails = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CPackageDetail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.PackageDeTails.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If

   QueryPackage = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPackage = False
End Function

Public Function DeletePackage(Ua As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeletePackage"
   DeletePackage = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePackage = True
End Function

Public Function CopyPackage(Ug As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyPackage"
   CopyPackage = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryPackage(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.ShowMode = SHOW_ADD
   Call Ug.SetFieldValue("PACKAGE_ID", -1)
   Call Ug.SetFieldValue("PACKAGE_NO", "***" & Ug.GetFieldValue("PACKAGE_NO"))
   Call Ug.SetFieldValue("PACKAGE_MASTER_FLAG", "N")
   
   Dim Det As CPackageDetail
   For Each Det In Ug.PackageDeTails
      Det.Flag = "A"
      Det.ShowMode = SHOW_ADD
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditPackage(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyPackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyPackage = False
End Function
Public Sub CreateBillingTransferItems(BD As CBillingDoc)
Dim Ti As CCashTransferItem
Dim TempCt As CCashTran
Dim IIEx As CCashTran
Dim Ei As CCashTran
Dim II As CCashTran
Dim I As Long
Dim j As Long
Dim AddFlag As Boolean
Dim AddFlag1 As Boolean
Dim AddFlag2 As Boolean
Dim AddFlag3 As Boolean
   
   Set BD.TransferItems = Nothing
   Set BD.TransferItems = New Collection
   
   For I = 1 To BD.Payments.Count
      Set TempCt = BD.Payments(I)
      If TempCt.GetFieldValue("STEP_ID") > 1 Then
         If TempCt.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And TempCt.GetFieldValue("TX_TYPE") = "I" Then
            Set IIEx = BD.Payments(I)
            AddFlag1 = True
         ElseIf TempCt.GetFieldValue("PAYMENT_TYPE") = CHEQUE_BANK_PMT And TempCt.GetFieldValue("TX_TYPE") = "E" Then
            Set Ei = BD.Payments(I)
            AddFlag2 = True
         ElseIf TempCt.GetFieldValue("PAYMENT_TYPE") = CASH_PMT Then
            Set II = BD.Payments(I)
            AddFlag3 = True
         End If
      Else
         j = I
         Set IIEx = BD.Payments(I)
         AddFlag = True
      End If
      If AddFlag Or (AddFlag1 And AddFlag2 And AddFlag3) Then
         Set Ti = New CCashTransferItem
         Set Ti.ImportItemEx = IIEx
         If (AddFlag1 And AddFlag2 And AddFlag3) Then
            Set Ti.ExportItem = Ei
            Set Ti.ImportItem = II
         End If
         Ti.Flag = "I"
         Call BD.TransferItems.add(Ti)
         Set Ti = Nothing
         AddFlag = False
         AddFlag1 = False
         AddFlag2 = False
         AddFlag3 = False
      End If
   Next I
End Sub
Public Function DeleteCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCashDoc"
   DeleteCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCashDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCashDoc = True
End Function
Public Function QueryCashDoc(Ua As CCashDoc, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCashDoc"
   QueryCashDoc = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.CashTranItems = Nothing
      Set Ua.CashTranItems = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHEQUE_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
            Call Ct.Cheque.SetFieldValue("TEMP_FEE_AMOUNT", Ct.GetFieldValue("FEE_AMOUNT"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.CashTranItems.add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim CDP As CCashDocPost
      Set CDP = New CCashDocPost
      Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing

      Set Ua.PostItems = Nothing
      Set Ua.PostItems = New Collection

      While Not m_Rs1.EOF
         Set CDP = New CCashDocPost
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         Call Ua.PostItems.add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend

   End If
   
   QueryCashDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCashDoc = False
End Function
Public Function AddEditCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim DocType As CASH_DOC_TYPE

   RName = "AddEditCashDoc"
   AddEditCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Ct As CCashTran
   Dim PaymentType As Long
   Dim OldPaymentType As Long
   
   DocType = Ua.GetFieldValue("DOCUMENT_TYPE")
   
   For Each Ct In Ua.CashTranItems
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.SetFieldValue("TX_NO", Ua.GetFieldValue("DOCUMENT_NO"))
      Call Ct.SetFieldValue("TX_DATE", Ua.GetFieldValue("DOCUMENT_DATE"))
      Call Ct.SetFieldValue("APAR_MAS_ID", Ua.GetFieldValue("APAR_MAS_ID"))
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If DocType = CASH_DEPOSIT Then 'ฝากเช็คที่มีอยู่ในระบบแล้ว
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHEQUE_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "Y")
            Call Ct.Cheque.SetFieldValue("POST_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("PAYMENT_TYPE") = CHEQUE_HAND_PMT Then 'เช็ค
               Ct.Cheque.ShowMode = SHOW_ADD
               Call Ct.Cheque.AddEditData
            End If
            Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
         End If
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'เมื่อเป็นการนำฝากจะไม่มีการแก้ไขเกิดขึ้น จะต้องลบแล้วสร้างใหม่
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT) Then
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHEQUE_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = CHEQUE_HAND_PMT) Then
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", -1)
            Call Ct.AddEditData
         
            Call Ct.Cheque.DeleteData
         ElseIf (PaymentType = OldPaymentType) And (PaymentType = CHEQUE_HAND_PMT) Then
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If DocType = CASH_DEPOSIT Then
            'ถ้าเป็นการลบ item สำหรับการนำเช็คในมือไปเข้า ธนาคาร ไม่ต้องลบเช็คนั้นทิ้งเพียงแต่ set BANK_FLAG ให้เป็น 'N'
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("OLD_CHEQUE_ID") > 0 Then 'เช็ค
               Call Ct.Cheque.DeleteData
            End If
         End If
      End If
   Next Ct
   
   For Each Ct In Ua.CashTranItems
        If (DocType = CASH_DEPOSIT) Then
          If Ct.GetFieldValue("TX_TYPE") = "I" Then
            Call Ct.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.SetFieldValue("BANK_ACCOUNT", Ua.GetFieldValue("BANK_ACCOUNT"))
            Call Ct.UpdateBankBranchAccount
         End If
      End If
   Next Ct
   
   Dim CDP As CCashDocPost
   Dim Ch As CCheque
   Set Ch = New CCheque
   For Each CDP In Ua.PostItems
      If CDP.Flag = "A" Then
         Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
         CDP.ShowMode = SHOW_ADD
         Call CDP.AddEditData
            
         Call Ch.SetFieldValue("CHEQUE_ID", CDP.GetFieldValue("CHEQUE_ID"))
         If CDP.GetFieldValue("POST_TYPE") = POST_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         ElseIf CDP.GetFieldValue("POST_TYPE") = WAITING_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "N")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         ElseIf CDP.GetFieldValue("POST_TYPE") = PASSED_CLEAR Then
            Call Ch.SetFieldValue("BANK_FLAG", "Y")
            Call Ch.SetFieldValue("POST_FLAG", "Y")
         End If
         
         Call Ch.UpdateBankFlag
         
      ElseIf CDP.Flag = "E" Then
         CDP.ShowMode = SHOW_EDIT
         Call CDP.AddEditData
      ElseIf CDP.Flag = "D" Then
         Call CDP.DeleteData
      End If
   Next CDP
   


   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCashDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCashDoc = False
End Function
Public Sub CreateCashTransferItems(Ivd As CCashDoc)
Dim Ti As CCashTransferItem
Dim O As Object
Dim Ei As CCashTran
Dim II As CCashTran
Dim I As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.CashTranItems.Count \ 2
   Count2 = Ivd.CashTranItems.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = Ivd.CashTranItems(j)
         If (O.GetFieldValue("TX_TYPE") = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.GetFieldValue("TX_TYPE") = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set Ei = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CCashTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = Ei
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.add(Ti)
      Set Ti = Nothing
   Next I
End Sub
Public Function AddEditBalanceVerify(Ua As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim I As Long
   RName = "AddEditBalanceVerify"
   AddEditBalanceVerify = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Gr As CBalanceVerifyDeTail
   For Each Gr In Ua.BalanceVerifyDetail
      Gr.BALANCE_VERIFY_ID = Ua.BALANCE_VERIFY_ID
      If Gr.Flag = "A" Then
         Gr.AddEditMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.AddEditMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
      
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   HasBegin = False
   IsOK = True

   AddEditBalanceVerify = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBalanceVerify = False
End Function

Public Function QueryBalanceVerify(Ua As CBalanceVerify, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBalanceVerify"
   QueryBalanceVerify = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CBalanceVerifyDeTail
      Set Gr = New CBalanceVerifyDeTail
      Gr.BALANCE_VERIFY_ID = Ua.BALANCE_VERIFY_ID
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.BalanceVerifyDetail = Nothing
      Set Ua.BalanceVerifyDetail = New Collection
      
      While Not m_Rs1.EOF
         Set Gr = New CBalanceVerifyDeTail
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.BalanceVerifyDetail.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If

   QueryBalanceVerify = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBalanceVerify = False
End Function

Public Function DeleteBalanceVerify(Ua As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

   RName = "DeleteBalanceVerify"
   DeleteBalanceVerify = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call Ua.DeleteData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBalanceVerify = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBalanceVerify = True
End Function
Public Function CopyBalanceVerify(Ug As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyBalanceVerify"
   CopyBalanceVerify = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryBalanceVerify(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.BALANCE_VERIFY_ID = -1
   Ug.BALANCE_VERIFY_NO = "***" & Ug.BALANCE_VERIFY_NO
   
   Dim Det As CBalanceVerifyDeTail
   For Each Det In Ug.BalanceVerifyDetail
      Det.Flag = "A"
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditBalanceVerify(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyBalanceVerify = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyBalanceVerify = False
End Function
Public Function CopyBalanceVerifyToAdjust(Ug As CBalanceVerify, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long
Dim Ivd As CInventoryDoc
Dim Lt As CLotItem
Dim InventoryBals2 As Collection
Dim TempLt As CLotItem
Dim Pt As CStockCode
Dim PartItemColl As Collection
   
   RName = "CopyBalanceVerifyToAdjust"
   CopyBalanceVerifyToAdjust = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryBalanceVerify(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Set InventoryBals2 = New Collection
   Set PartItemColl = New Collection
   
   Call LoadLeftAmountLotItem(InventoryBals2, -1, Ug.BALANCE_VERIFY_DATE)
   Call LoadStockCode(Nothing, PartItemColl)
      
   Set Ivd = New CInventoryDoc
   Ivd.ShowMode = SHOW_ADD
   Ivd.INVENTORY_DOC_ID = -1
   Ivd.DOCUMENT_NO = Ug.BALANCE_VERIFY_NO & " / " & Now
   Ivd.DOCUMENT_DATE = Ug.BALANCE_VERIFY_DATE
   Ivd.DOCUMENT_TYPE = ADJUST_DOCTYPE
   
   Ivd.DOCUMENT_DESC = "มากจากการตั้งค่าของหน้าจอตรวจยอดสต็อค" & Now
   Ivd.CANCEL_FLAG = "N"
   
   Dim Det As CBalanceVerifyDeTail
   For Each Det In Ug.BalanceVerifyDetail
      Set TempLt = GetObject("CLotItem", InventoryBals2, Trim(Det.LOCATION_ID & "-" & Det.PART_ITEM_ID))
      If Not (TempLt.SUM_AMOUNT = Det.PART_ITEM_AMOUNT) Then
         Set Lt = New CLotItem
         Lt.Flag = "A"
         If TempLt.SUM_AMOUNT > Det.PART_ITEM_AMOUNT Then    ' ต้องปรับ ออก
            Lt.TX_AMOUNT = TempLt.SUM_AMOUNT - Det.PART_ITEM_AMOUNT
            Lt.TX_TYPE = "E"
            Lt.MULTIPLIER = -1
         Else
            Lt.TX_AMOUNT = Det.PART_ITEM_AMOUNT - TempLt.SUM_AMOUNT
            Lt.TX_TYPE = "I"
            Lt.MULTIPLIER = 1
         End If
                           
         Lt.PART_ITEM_ID = Det.PART_ITEM_ID
         Lt.LOCATION_ID = Det.LOCATION_ID
         
         Set Pt = GetObject("CStockCode", PartItemColl, Trim(Str(Det.PART_ITEM_ID)))
         Lt.UNIT_TRAN_ID = Pt.UNIT_CHANGE_ID
         Lt.UNIT_MULTIPLE = 1
         
         Call Ivd.ImportExportItems.add(Lt)
         Set Lt = Nothing
      End If
   Next Det
      
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditInventoryDoc(Ivd, IsOK, False, ErrorObj)
   
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   Set InventoryBals2 = Nothing
   Set Ivd = Nothing
   Set PartItemColl = Nothing
   
   HasBegin = False
   IsOK = True

   CopyBalanceVerifyToAdjust = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyBalanceVerifyToAdjust = False
   Set InventoryBals2 = Nothing
   Set Ivd = Nothing
End Function

Public Function QueryLockDate(Ua As CLockDate, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryLockDate"
   QueryLockDate = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   QueryLockDate = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLockDate = False
End Function
Public Function AddEditLockDate(Ua As CLockDate, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLockDate"
   AddEditLockDate = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditLockDate = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLockDate = False
End Function

Public Function CopyStockCode(Ug As CStockCode, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long

   RName = "CopyStockCode"
   CopyStockCode = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryStockCode(Ug, TempRs, iCount, IsOK, ErrorObj)
   
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.ShowMode = SHOW_ADD
   Ug.STOCK_CODE_ID = -1
   Ug.STOCK_NO = Ug.STOCK_NO & "*"
         
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditStockCode(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyStockCode = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyStockCode = False
End Function

Public Function DeleteJob(Ua As CJob, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
   
   RName = "DeleteJob"
   DeleteJob = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ua.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteJob = True
End Function
Public Function QueryJob(Ua As CJob, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryJob"
   QueryJob = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      
      Dim Ji As CJobItem
      Set Ji = New CJobItem
      Ji.JOB_ID = Ua.JOB_ID
      Call Ji.QueryData(1, m_Rs1, iCount, True)
      Set Ji = Nothing
      
      Set Ua.CollJobInputs = Nothing
      Set Ua.CollJobInputs = New Collection
      Set Ua.CollJobOutputs = Nothing
      Set Ua.CollJobOutputs = New Collection
      
      While Not m_Rs1.EOF
         Set Ji = New CJobItem
         Call Ji.PopulateFromRS(1, m_Rs1)
         
         If Ji.TX_TYPE = "I" Then      'ด้าน Job เข้าแสดงว่าต้องเบิกของออกมาจากคลัง
            Call Ua.CollJobInputs.add(Ji)
         Else
            Call Ua.CollJobOutputs.add(Ji)
         End If
         
         Set Ji = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryJob = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryJob = False
End Function
Public Function AddEditJob(Ua As CJob, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditJob"
   AddEditJob = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
      
   Dim Cd As CConfigDoc
   If Ua.AddEditMode = SHOW_ADD Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ua.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ua.JOB_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ua.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim Ji As CJobItem
   
   For Each Ji In Ua.CollJobInputs
      If Ji.Flag = "A" Then
         Ji.JOB_ID = Ua.JOB_ID
         Ji.AddEditMode = SHOW_ADD
         Call Ji.AddEditData
      ElseIf Ji.Flag = "E" Then
         Ji.AddEditMode = SHOW_EDIT
         Call Ji.AddEditData
      ElseIf Ji.Flag = "D" Then
         Call Ji.DeleteData
      End If
   Next Ji
   
   For Each Ji In Ua.CollJobOutputs
      If Ji.Flag = "A" Then
         Ji.JOB_ID = Ua.JOB_ID
         Ji.AddEditMode = SHOW_ADD
         Call Ji.AddEditData
      ElseIf Ji.Flag = "E" Then
         Ji.AddEditMode = SHOW_EDIT
         Call Ji.AddEditData
      ElseIf Ji.Flag = "D" Then
         Call Ji.DeleteData
      End If
   Next Ji
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditJob = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.Description
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditJob = False
End Function

Public Function Job2InventoryDoc(Jb As CJob, Ivd As CInventoryDoc, Optional DocumentType As Long = -1) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Ji As CJobItem
Dim Ei As CLotItem
   Set Ivd = Nothing
   
   Set Ivd = New CInventoryDoc
   
   If Jb.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = Jb.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      Call Ivd.PopulateFromRS(1, TempRs)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.ShowMode = SHOW_EDIT
   Else
      Ivd.ShowMode = SHOW_ADD
   End If
   
   Ivd.DOCUMENT_DATE = Jb.JOB_DATE
   Ivd.DOCUMENT_NO = Jb.JOB_NO
   Ivd.DOCUMENT_TYPE = DocumentType
   
   Ivd.CANCEL_FLAG = "N"
   
   For Each Ji In Jb.CollJobInputs ' In ของ Job ก็ คืนการเบิกออกจากคลัง
      If Ji.Flag = "A" Then
         Set Ei = New CLotItem
               
         Ei.TX_TYPE = "E"
         Ei.MULTIPLIER = -1
            
         Ei.UNIT_TRAN_ID = Ji.UNIT_TRAN_ID
         Ei.UNIT_MULTIPLE = Ji.UNIT_MULTIPLE
         
         Ei.Flag = "A"
         
         Ei.PART_ITEM_ID = Ji.PART_ITEM_ID
         Ei.LOCATION_ID = Ji.LOCATION_ID
         
         Ei.TX_AMOUNT = Ji.TX_AMOUNT
         
         Ei.LINK_ID = Ji.LINK_ID
         Ei.CANCEL_FLAG = "N"
                        
         Call Ivd.ImportExportItems.add(Ei)
         Set Ei = Nothing
      ElseIf Ji.Flag = "D" Then
         Set Ei = GetExportItem(Ivd, Ji.LINK_ID)
         Ei.Flag = "D"
      End If
   Next Ji
   
   For Each Ji In Jb.CollJobOutputs ' Out ของ Job ก็ คืนการเก็บสินค้าที่ได้เข้าคลังหลังจากผลิตเสร็จ
      If Ji.LOCATION_ID > 0 Then
         If Ji.Flag = "A" Then
            Set Ei = New CLotItem
            
            Ei.TX_TYPE = "I"
            Ei.MULTIPLIER = 1
            
            Ei.UNIT_TRAN_ID = Ji.UNIT_TRAN_ID
            Ei.UNIT_MULTIPLE = Ji.UNIT_MULTIPLE
            
            Ei.Flag = "A"
            Ei.PART_ITEM_ID = Ji.PART_ITEM_ID
            
            Ei.LOCATION_ID = Ji.LOCATION_ID
            Ei.TX_AMOUNT = Ji.TX_AMOUNT
            
            Ei.LINK_ID = Ji.LINK_ID
            Ei.CANCEL_FLAG = "N"
            
            Call Ivd.ImportExportItems.add(Ei)
            
            Set Ei = Nothing
         ElseIf Ji.Flag = "D" Then
            Set Ei = GetExportItem(Ivd, Ji.LINK_ID)
            Ei.Flag = "D"
         End If
      End If
   Next Ji
   
End Function

