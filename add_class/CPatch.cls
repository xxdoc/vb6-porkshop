VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Conn As ADODB.Connection
Private m_Conn2 As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set m_Conn2 = glbDatabaseMngr.DB2Connection
End Sub
Public Function IsPatch(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long
Dim Result As Boolean
   
   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERSION = '" & VersionStr & "' "
   Call glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, TempRs, glbErrorLog)
   
   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch = Result
   Exit Function
   
ErrorHandler:
   IsPatch = False
End Function
'Public Function Patch_2014_07_07_1_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_07_1_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE NETWORK_CHART ADD HOLD_FLAG FLAG_TYPE DEFAULT 'N' "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_07_1_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_07_2_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_07_2_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "UPDATE NETWORK_CHART SET HOLD_FLAG = 'N' "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_07_2_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_13_1_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_1_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "CREATE TABLE  FORMULA  "
'   SQL = SQL & "( "
'   SQL = SQL & "FORMULA_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "FORMULA_NO                    CODE_TYPE , "
'   SQL = SQL & "FORMULA_DATE                    DATE_TYPE , "
'   SQL = SQL & "FORMULA_DESC                    DESC_TYPE , "
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE FORMULA ADD CONSTRAINT FORMULA_PK  PRIMARY KEY (FORMULA_ID) "
'   m_Conn.Execute (SQL)
'
'   SQL = "CREATE GENERATOR FORMULA_SEQ "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_1_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_13_2_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_2_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "                 CREATE TRIGGER FORMULA_BA FOR FORMULA "
'   SQL = SQL & "  ACTIVE BEFORE INSERT "
'   SQL = SQL & "  POSITION 0 "
'   SQL = SQL & "  AS "
'   SQL = SQL & "  BEGIN "
'   SQL = SQL & "        NEW.FORMULA_ID  = GEN_ID (FORMULA_SEQ, 1); "
'   SQL = SQL & "  END; "
'
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_2_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_13_3_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_3_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "CREATE TABLE  FORMULA_ITEM  "
'   SQL = SQL & "( "
'   SQL = SQL & "FORMULA_ITEM_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "FORMULA_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "PART_ITEM_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "LOCATION_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "TX_TYPE                    CODE_TYPE , "
'   SQL = SQL & "TX_AMOUNT                    MONEY_TYPE , "
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_PK  PRIMARY KEY (FORMULA_ITEM_ID) "
'   m_Conn.Execute (SQL)
'
'   SQL = "CREATE GENERATOR FORMULA_ITEM_SEQ "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_3_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'
'Public Function Patch_2014_07_13_4_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_4_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "                 CREATE TRIGGER FORMULA_ITEM_BA FOR FORMULA_ITEM "
'   SQL = SQL & "  ACTIVE BEFORE INSERT "
'   SQL = SQL & "  POSITION 0 "
'   SQL = SQL & "  AS "
'   SQL = SQL & "  BEGIN "
'   SQL = SQL & "        NEW.FORMULA_ITEM_ID  = GEN_ID (FORMULA_ITEM_SEQ, 1); "
'   SQL = SQL & "  END; "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_4_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'
'Public Function Patch_2014_07_13_5_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_5_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_FK FOREIGN KEY (FORMULA_ID) REFERENCES FORMULA "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_PART_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES STOCK_CODE "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_LC_FK FOREIGN KEY (LOCATION_ID) REFERENCES MASTER_REF "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_5_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_13_6_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_6_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE FORMULA_ITEM ADD UNIT_TRAN_ID ID_TYPE "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE FORMULA_ITEM ADD UNIT_MULTIPLE MONEY_TYPE "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_6_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_13_7_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_13_7_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_UNIT_FK FOREIGN KEY (UNIT_TRAN_ID) REFERENCES MASTER_REF "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_13_7_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_21_1_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_21_1_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "CREATE TABLE  JOB "
'   SQL = SQL & "( "
'   SQL = SQL & "JOB_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "JOB_NO                    CODE_TYPE , "
'   SQL = SQL & "JOB_DATE                    DATE_TYPE , "
'   SQL = SQL & "JOB_DESC                    DESC_TYPE , "
'   SQL = SQL & "FORMULA_ID             ID_TYPE NOT NULL, "              'ผลิตจากสูตรอะไร
'   SQL = SQL & "FORMULA_AMOUNT             MONEY_TYPE, "              'ใช้สูตรจำนวนเท่าไหร่
'   SQL = SQL & "INVENTORY_DOC_ID             ID_TYPE NOT NULL, "              'LINK STOCK
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE JOB ADD CONSTRAINT JOB_PK  PRIMARY KEY (JOB_ID) "
'   m_Conn.Execute (SQL)
'
'   SQL = "CREATE GENERATOR JOB_SEQ "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_21_1_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_21_2_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_21_2_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "                 CREATE TRIGGER JOB_BA FOR JOB "
'   SQL = SQL & "  ACTIVE BEFORE INSERT "
'   SQL = SQL & "  POSITION 0 "
'   SQL = SQL & "  AS "
'   SQL = SQL & "  BEGIN "
'   SQL = SQL & "        NEW.JOB_ID  = GEN_ID (JOB_SEQ, 1); "
'   SQL = SQL & "  END; "
'
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_21_2_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_21_3_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_21_3_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE JOB ADD CONSTRAINT JOB_FORMULA_FK FOREIGN KEY (FORMULA_ID) REFERENCES FORMULA "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_21_3_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_21_4_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_21_4_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE JOB ADD CONSTRAINT JOB_INVENTORY_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_21_4_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_23_1_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_23_1_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "CREATE TABLE  JOB_ITEM  "
'   SQL = SQL & "( "
'   SQL = SQL & "JOB_ITEM_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "JOB_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "PART_ITEM_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "LOCATION_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "LINK_ID             ID_TYPE NOT NULL, "
'   SQL = SQL & "TX_TYPE                    CODE_TYPE , "
'   SQL = SQL & "TX_AMOUNT                    MONEY_TYPE , "
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE JOB_ITEM ADD CONSTRAINT JOB_ITEM_PK  PRIMARY KEY (JOB_ITEM_ID) "
'   m_Conn.Execute (SQL)
'
'   SQL = "CREATE GENERATOR JOB_ITEM_SEQ "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_23_1_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'
'Public Function Patch_2014_07_23_2_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_23_2_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "                 CREATE TRIGGER JOB_ITEM_BA FOR JOB_ITEM "
'   SQL = SQL & "  ACTIVE BEFORE INSERT "
'   SQL = SQL & "  POSITION 0 "
'   SQL = SQL & "  AS "
'   SQL = SQL & "  BEGIN "
'   SQL = SQL & "        NEW.JOB_ITEM_ID  = GEN_ID (JOB_ITEM_SEQ, 1); "
'   SQL = SQL & "  END; "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_23_2_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'
'Public Function Patch_2014_07_23_3_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_23_3_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE JOB_ITEM ADD CONSTRAINT JOB_ITEM_FK FOREIGN KEY (JOB_ID) REFERENCES JOB "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE JOB_ITEM ADD CONSTRAINT JOB_ITEM_PART_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES STOCK_CODE "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE JOB_ITEM ADD CONSTRAINT JOB_ITEM_LC_FK FOREIGN KEY (LOCATION_ID) REFERENCES MASTER_REF "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_23_3_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_23_4_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_23_4_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE JOB_ITEM ADD UNIT_TRAN_ID ID_TYPE "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE JOB_ITEM ADD UNIT_MULTIPLE MONEY_TYPE "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_23_4_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2014_07_23_5_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'Dim PD As CPatchDelivery
'Set PD = New CPatchDelivery
'
'   VersionStr = "2014_07_23_5_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = " ALTER TABLE JOB_ITEM ADD CONSTRAINT JOB_ITEM_UNIT_FK FOREIGN KEY (UNIT_TRAN_ID) REFERENCES MASTER_REF "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   Call PD.SetFieldValue("PATCH_DELIVERY_ID", -1)
'   Call PD.SetFieldValue("PATCH_VERSION", VersionStr)
'   Call PD.SetFieldValue("PATCH_DATE", InternalDateToDate(ServerDateTime))
'   Call PD.SetFieldValue("PATCH_BY", glbUser.USER_ID)
'   Call PD.AddEditData(SHOW_ADD)
'
'   m_Conn.CommitTrans
'
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2014_07_23_5_jill = True
'   Exit Function
'
'ErrorHandler:
'
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
