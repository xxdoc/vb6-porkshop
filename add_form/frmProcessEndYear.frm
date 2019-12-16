VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcessEndYear 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmProcessEndYear.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6555
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   11562
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   3
         Top             =   1560
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblNote 
         Caption         =   "Label1"
         Height          =   2835
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   11655
      End
      Begin VB.Label lblProgressDesc 
         Caption         =   "Label1"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   2640
         Width           =   9255
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   1
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmProcessEndYear.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   9
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   2
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmProcessEndYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public HeaderText As String

Private m_Conn As ADODB.Connection

Private Sub cmdStart_Click()
Dim Status As Boolean
Dim IsOK As Boolean

   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
   
   'Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   Status = AdjustStockCode
   
   Me.Enabled = True
   
   If Status Then
      'If ConfirmSave Then
         'Call glbDaily.CommitTransaction
         glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
         glbErrorLog.ShowUserError
'      Else
'         Call glbDaily.RollbackTransaction
'         glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
'         glbErrorLog.ShowUserError
'      End If
   Else
'      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
      glbErrorLog.ShowUserError
   End If
   
   OKClick = True
   Unload Me
   Exit Sub
   
End Sub
Private Sub Form_Activate()
      Me.Refresh
      DoEvents
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   End If
End Sub
Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblToDate, "ถึงวันที่", RGB(255, 0, 0))
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call InitNormalLabel(lblProgressDesc, "กรอกวันที่ จากนั้น กดเริ่ม")
   Call InitNormalLabel(lblNote, "ทำเสร็จวันที่ 07/01/2557 Test เรียบร้อย เริ่มใช้งาน จันทร์ 13/01/2557 ( วันที่ม็อบสุเทพจะปิดกรุงเทพ มันบ้า) จะใช้งาน เริ่มตั้งแต่  ถึงวันที่ 31/12/2555" & vbCrLf & "โดยการประมวลผลสิ้นปีนี้มีเงื่อนไขดังนี้" & vbCrLf & "1.ควร COPY ออกมา Test ในเครื่อง StandAlone ก่อน" & vbCrLf & "2.BACKUP / ตรวจสอบยอดหนี้ Stock ก่อนหลังการประมวลผล")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Call EnableForm(Me, False)
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Function AdjustStockCode() As Boolean
Dim m_LotItem As CLotItem
Dim I As Long
Dim IsOK As Boolean
Dim TempPartColl As Collection
Dim InventoryBalBeforeAdjustColl As Collection
Dim InventoryBalAferAdjustColl As Collection
Dim TempDate As String
Dim SQL1 As String
Dim Ivd As CInventoryDoc
Dim Mr As CMasterRef
Dim Pi As CStockCode
Dim TempLotItem As CLotItem
Dim TempLotItemSearch As CLotItem
Dim Amt As Double
   
   Set TempPartColl = New Collection
   Set InventoryBalBeforeAdjustColl = New Collection
   Set InventoryBalAferAdjustColl = New Collection
   
   I = 0
   prgProgress.Min = 0
   prgProgress.Max = 100
   
   AdjustStockCode = False
   
   Call LoadStockCode(Nothing, TempPartColl)
   
   Call LoadLeftAmountLocation(InventoryBalBeforeAdjustColl, , uctlToDate.ShowDate)
   
   prgProgress.Value = 1
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   '/*REMOVE CONSTRAINTS*/
   '/*CREATE CONSTRAINTS CASCADE OR SET NULL*/
   lblProgressDesc.Caption = "กำลัง REMOVE CONSTRAINTS AND CREATE CONSTRAINTS CASCADE OR SET NULL"
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT BILLING_DOC_RCPCNDN_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT BILLING_DOC_RCPCNDN_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT BILLS_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT BILLS_ID_FK FOREIGN KEY (BILLS_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT DOC_ID_BILLS_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT DOC_ID_BILLS_ID_FK FOREIGN KEY (DOC_ID_BILLS) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 2
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT DOC_ID_FK FOREIGN KEY (DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT DOC_ID_RCP_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT DOC_ID_RCP_ID_FK FOREIGN KEY (DOC_ID_RCP) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE DOC_ITEM DROP CONSTRAINT DOC_ITEM_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM ADD CONSTRAINT DOC_ITEM_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 3
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE DOC_ITEM_LINK DROP CONSTRAINT DOC_ITEM_LINK_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM_LINK ADD CONSTRAINT DOC_ITEM_LINK_FK FOREIGN KEY (DOC_ITEM_ID) REFERENCES DOC_ITEM(DOC_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE PRINT_LABEL DROP CONSTRAINT DOC_ITEM_ID_LABEL_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE PRINT_LABEL ADD CONSTRAINT DOC_ITEM_ID_LABEL_FK FOREIGN KEY (DOC_ITEM_ID) REFERENCES DOC_ITEM(DOC_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 4
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE DOC_ITEM DROP CONSTRAINT DOC_ITEM_PO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM ADD CONSTRAINT DOC_ITEM_PO_ID_FK FOREIGN KEY (PO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE SET NULL;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE BILLING_DOC DROP CONSTRAINT BILLING_DOC_SR_REF_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_SR_REF_DO_ID_FK FOREIGN KEY (SR_REF_DO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE SET NULL;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 5
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE CASH_TRAN DROP CONSTRAINT CASH_TRAN_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE BILLING_ADDITION DROP CONSTRAINT BILLING_ADDITION_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_ADDITION ADD CONSTRAINT BILLING_ADDITION_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE BILLING_SUBTRACT DROP CONSTRAINT BILLING_SUBTRACT_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_SUBTRACT ADD CONSTRAINT BILLING_SUBTRACT_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 6
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE LOT_ITEM DROP CONSTRAINT LOT_ITEM_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM ADD CONSTRAINT LOT_ITEM_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE DOC_ITEM_LINK DROP CONSTRAINT DOCIMPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM_LINK ADD CONSTRAINT DOCIMPORT_LOT_ITEM_FK FOREIGN KEY (IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE DOC_ITEM_LINK DROP CONSTRAINT DOCMAIN_IMPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM_LINK ADD CONSTRAINT DOCMAIN_IMPORT_LOT_ITEM_FK FOREIGN KEY (MAIN_IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 8
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE BILLING_DOC DROP CONSTRAINT BILLING_DOC_INV_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_INV_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID) ON DELETE SET NULL;"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE LOT_ITEM_LINK DROP CONSTRAINT EXPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM_LINK ADD CONSTRAINT EXPORT_LOT_ITEM_FK FOREIGN KEY (EXPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   prgProgress.Value = 9
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE LOT_ITEM_LINK DROP CONSTRAINT IMPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM_LINK ADD CONSTRAINT IMPORT_LOT_ITEM_FK FOREIGN KEY (IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE LOT_ITEM_LINK DROP CONSTRAINT LOT_LINK_MAIN_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM_LINK ADD CONSTRAINT LOT_LINK_MAIN_ID_FK FOREIGN KEY (MAIN_IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID) ON DELETE CASCADE;"
   Call m_Conn.Execute(SQL1)
   
   
   '/*REMOVE CONSTRAINTS*/
   '/*CREATE CONSTRAINTS CASCADE OR SET NULL*/
   lblProgressDesc.Caption = "ลบเอกสารเอกสารนำส่งยอดใบวางบิล"
   prgProgress.Value = 10
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   TempDate = DateToStringIntHi(uctlToDate.ShowDate)
   
   '/*BILL DETAIL*/
   SQL1 = "DELETE FROM BILL_DETAIL BDT WHERE BDT.SUM_BILL_ID IN (SELECT BD.SUM_BILL_ID FROM SUM_BILL BD WHERE BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   Call m_Conn.Execute(SQL1)
      
   '/*SUM BILL*/
   SQL1 = "DELETE FROM SUM_BILL BD WHERE BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบลดหนี้เพิ่มหนี้"
   prgProgress.Value = 15
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*CN DN*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 7 OR BD.DOCUMENT_TYPE = 8) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & "AND BD.PAY_AMOUNT ="
   SQL1 = SQL1 & " " & "("
   SQL1 = SQL1 & " " & "SELECT SUM(EN.PAID_AMOUNT) PAID_AMOUNT FROM RCPCNDN_ITEM EN"
   SQL1 = SQL1 & " " & "LEFT OUTER JOIN BILLING_DOC BDRE ON (BDRE.BILLING_DOC_ID = EN.BILLING_DOC_ID)"
   SQL1 = SQL1 & " " & "WHERE BDRE.DOCUMENT_TYPE = 5 AND EN.DOC_ID = BD.BILLING_DOC_ID AND BDRE.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & ")"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบรับคืนสินค้า"
   prgProgress.Value = 20
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

   '/*SR*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 6) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & "AND (BD.TOTAL_PRICE+BD.VAT_AMOUNT-BD.DISCOUNT_AMOUNT-BD.EXT_DISCOUNT_AMOUNT) ="
   SQL1 = SQL1 & " " & "("
   SQL1 = SQL1 & " " & "SELECT SUM(EN.PAID_AMOUNT) PAID_AMOUNT FROM RCPCNDN_ITEM EN"
   SQL1 = SQL1 & " " & "LEFT OUTER JOIN BILLING_DOC BDRE ON (BDRE.BILLING_DOC_ID = EN.BILLING_DOC_ID)"
   SQL1 = SQL1 & " " & "WHERE BDRE.DOCUMENT_TYPE = 5 AND EN.DOC_ID = BD.BILLING_DOC_ID AND BDRE.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & ")"
   Call m_Conn.Execute(SQL1)
  
   lblProgressDesc.Caption = "ลบเอกสารใบขายสด"
   prgProgress.Value = 25
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
  
   '/*KAY SOD*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 4) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบขายเชื่อ"
   prgProgress.Value = 27
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*INVOICE*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 3) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & "AND (BD.TOTAL_PRICE+BD.VAT_AMOUNT-BD.DISCOUNT_AMOUNT-BD.EXT_DISCOUNT_AMOUNT) ="
   SQL1 = SQL1 & " " & "("
   SQL1 = SQL1 & " " & "SELECT SUM(EN.PAID_AMOUNT) PAID_AMOUNT FROM RCPCNDN_ITEM EN"
   SQL1 = SQL1 & " " & "LEFT OUTER JOIN BILLING_DOC BDRE ON (BDRE.BILLING_DOC_ID = EN.BILLING_DOC_ID)"
   SQL1 = SQL1 & " " & "WHERE BDRE.DOCUMENT_TYPE = 5 AND EN.DOC_ID = BD.BILLING_DOC_ID AND BDRE.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & ")"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบ PO"
   prgProgress.Value = 30
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*PO*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 2) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบเสร็จรับชำระ"
   prgProgress.Value = 35
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*RECEIPT*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 5 OR BD.DOCUMENT_TYPE = 9) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & "AND ((SELECT COUNT(*) FROM RCPCNDN_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0);"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบเสร็จรับชำระ (เป็นชุด)"
   prgProgress.Value = 40
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*RECEIPT PACK*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 10) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   SQL1 = SQL1 & " " & "AND ((SELECT COUNT(*) FROM RCPCNDN_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0);"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบ STOCK ที่อ้างอิงจากใบขายเชื่อ ขายสด รับคืน"
   prgProgress.Value = 45
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*IVTRD FROM INVOICE KAY SOD SR*/
   SQL1 = "DELETE FROM INVENTORY_DOC IVTRD"
   SQL1 = SQL1 & " " & "WHERE (IVTRD.DOCUMENT_TYPE = 10 Or IVTRD.DOCUMENT_TYPE = 21 Or IVTRD.DOCUMENT_TYPE = 30) AND IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบ GR RR RO"
   prgProgress.Value = 50
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*GR*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 106) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
      
   '/*RR*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 103) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*RO*/
   SQL1 = "DELETE FROM BILLING_DOC BD"
   SQL1 = SQL1 & " " & "WHERE (BD.DOCUMENT_TYPE = 102) AND BD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*IVTRD FROM RR GR*/
   SQL1 = "DELETE FROM INVENTORY_DOC IVTRD"
   SQL1 = SQL1 & " " & "WHERE (IVTRD.DOCUMENT_TYPE = 11 Or IVTRD.DOCUMENT_TYPE = 22 Or IVTRD.DOCUMENT_TYPE = 31) AND IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "ลบเอกสารใบ JOB"
   prgProgress.Value = 55
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*JOB ITEM*/
   SQL1 = "DELETE FROM JOB_ITEM JBIT"
   SQL1 = SQL1 & " " & "WHERE JBIT.JOB_ID"
   SQL1 = SQL1 & " " & "IN (SELECT JB.JOB_ID FROM JOB JB WHERE JB.JOB_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   Call m_Conn.Execute(SQL1)
   
   '/*JOB*/
   SQL1 = "DELETE FROM JOB JB"
   SQL1 = SQL1 & " " & "WHERE JB.JOB_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*IVTRD FROM JOB*/
   SQL1 = "DELETE FROM INVENTORY_DOC IVTRD"
   SQL1 = SQL1 & " " & "WHERE (IVTRD.DOCUMENT_TYPE = 1000) AND IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*BALANCE_ACCUM NO USE NOW*/
   SQL1 = "DELETE FROM BALANCE_ACCUM BLAC;"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "เอกสารนับยอด STOCK และ เป้าการขาย"
   prgProgress.Value = 60
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents

   '/*BALANCE VERIFY DETAIL*/
   SQL1 = "DELETE FROM BALANCE_VERIFY_DETAIL BLVRFDT"
   SQL1 = SQL1 & " " & "WHERE BLVRFDT.BALANCE_VERIFY_ID"
   SQL1 = SQL1 & " " & "IN (SELECT BLVRF.BALANCE_VERIFY_ID FROM BALANCE_VERIFY BLVRF WHERE BLVRF.BALANCE_VERIFY_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   Call m_Conn.Execute(SQL1)
   
   '/*BALANCE VERIFY*/
   SQL1 = "DELETE FROM BALANCE_VERIFY BLVRF WHERE BLVRF.BALANCE_VERIFY_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*TAGET DETAIL*/
   SQL1 = "DELETE FROM TAGET_DETAIL TGDT"
   SQL1 = SQL1 & " " & "WHERE TGDT.TAGET_ID"
   SQL1 = SQL1 & " " & "IN (SELECT TG.TAGET_ID FROM TAGET TG WHERE TG.YYYYMM <='" & Year(uctlToDate.ShowDate) & Format(Month(uctlToDate.ShowDate), "00") & "')"
   Call m_Conn.Execute(SQL1)

   '/*TAGET*/
   SQL1 = "DELETE FROM TAGET TG WHERE TG.YYYYMM <= '" & Year(uctlToDate.ShowDate) & Format(Month(uctlToDate.ShowDate), "00") & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "เอกสารใบรับเข้า เบิกออก โอนย้าย ปรับยอด คลัง"
   prgProgress.Value = 65
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*IVTRD IMPORT EXPORT*/
   SQL1 = "DELETE FROM INVENTORY_DOC IVTRD"
   SQL1 = SQL1 & " " & "WHERE (IVTRD.DOCUMENT_TYPE = 1 Or IVTRD.DOCUMENT_TYPE = 2 Or IVTRD.DOCUMENT_TYPE = 3 Or IVTRD.DOCUMENT_TYPE = 4) AND IVTRD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "LOCK เอกสาร"
   prgProgress.Value = 70
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   '/*COMMIT BILLING_DOC*/
   SQL1 = "UPDATE BILLING_DOC SET COMMIT_FLAG = 'Y' WHERE DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*COMMIT INVENTORY_DOC*/
   SQL1 = "UPDATE INVENTORY_DOC SET COMMIT_FLAG = 'Y' WHERE DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "'"
   Call m_Conn.Execute(SQL1)
   
   '/*REMOVE CONSTRAINTS*/
   '/*CREATE CONSTRAINTS*/
   lblProgressDesc.Caption = "กำลัง REMOVE CONSTRAINTS AND CREATE CONSTRAINTS"
   prgProgress.Value = 75
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT BILLING_DOC_RCPCNDN_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT BILLING_DOC_RCPCNDN_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT BILLS_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT BILLS_ID_FK FOREIGN KEY (BILLS_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT DOC_ID_BILLS_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT DOC_ID_BILLS_ID_FK FOREIGN KEY (DOC_ID_BILLS) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT DOC_ID_FK FOREIGN KEY (DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE RCPCNDN_ITEM DROP CONSTRAINT DOC_ID_RCP_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE RCPCNDN_ITEM ADD CONSTRAINT DOC_ID_RCP_ID_FK FOREIGN KEY (DOC_ID_RCP) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE DOC_ITEM DROP CONSTRAINT DOC_ITEM_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM ADD CONSTRAINT DOC_ITEM_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE DOC_ITEM_LINK DROP CONSTRAINT DOC_ITEM_LINK_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM_LINK ADD CONSTRAINT DOC_ITEM_LINK_FK FOREIGN KEY (DOC_ITEM_ID) REFERENCES DOC_ITEM(DOC_ITEM_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE PRINT_LABEL DROP CONSTRAINT DOC_ITEM_ID_LABEL_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE PRINT_LABEL ADD CONSTRAINT DOC_ITEM_ID_LABEL_FK FOREIGN KEY (DOC_ITEM_ID) REFERENCES DOC_ITEM(DOC_ITEM_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE DOC_ITEM DROP CONSTRAINT DOC_ITEM_PO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM ADD CONSTRAINT DOC_ITEM_PO_ID_FK FOREIGN KEY (PO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE BILLING_DOC DROP CONSTRAINT BILLING_DOC_SR_REF_DO_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_SR_REF_DO_ID_FK FOREIGN KEY (SR_REF_DO_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE CASH_TRAN DROP CONSTRAINT CASH_TRAN_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE BILLING_ADDITION DROP CONSTRAINT BILLING_ADDITION_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_ADDITION ADD CONSTRAINT BILLING_ADDITION_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE BILLING_SUBTRACT DROP CONSTRAINT BILLING_SUBTRACT_BILLING_DOC_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_SUBTRACT ADD CONSTRAINT BILLING_SUBTRACT_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE LOT_ITEM DROP CONSTRAINT LOT_ITEM_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM ADD CONSTRAINT LOT_ITEM_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE DOC_ITEM_LINK DROP CONSTRAINT DOCIMPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM_LINK ADD CONSTRAINT DOCIMPORT_LOT_ITEM_FK FOREIGN KEY (IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE DOC_ITEM_LINK DROP CONSTRAINT DOCMAIN_IMPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE DOC_ITEM_LINK ADD CONSTRAINT DOCMAIN_IMPORT_LOT_ITEM_FK FOREIGN KEY (MAIN_IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE BILLING_DOC DROP CONSTRAINT BILLING_DOC_INV_DOC_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_INV_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID);"
   Call m_Conn.Execute(SQL1)

   SQL1 = "ALTER TABLE LOT_ITEM_LINK DROP CONSTRAINT EXPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM_LINK ADD CONSTRAINT EXPORT_LOT_ITEM_FK FOREIGN KEY (EXPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE LOT_ITEM_LINK DROP CONSTRAINT IMPORT_LOT_ITEM_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM_LINK ADD CONSTRAINT IMPORT_LOT_ITEM_FK FOREIGN KEY (IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID);"
   Call m_Conn.Execute(SQL1)
   
   SQL1 = "ALTER TABLE LOT_ITEM_LINK DROP CONSTRAINT LOT_LINK_MAIN_ID_FK;"
   Call m_Conn.Execute(SQL1)
   SQL1 = "ALTER TABLE LOT_ITEM_LINK ADD CONSTRAINT LOT_LINK_MAIN_ID_FK FOREIGN KEY (MAIN_IMPORT_LOT_ITEM_ID) REFERENCES LOT_ITEM(LOT_ITEM_ID);"
   Call m_Conn.Execute(SQL1)
   
   lblProgressDesc.Caption = "กำลัง ปรับยอด STOCK"
   prgProgress.Value = 80
   txtPercent.Text = prgProgress.Value
   Me.Refresh
   DoEvents
   
   'LOAD STOCK BALANCE AFTER ADJUST END YEAR
   Call LoadLeftAmountLocation(InventoryBalAferAdjustColl, , uctlToDate.ShowDate)
   
   Set Ivd = New CInventoryDoc
   Ivd.ShowMode = SHOW_ADD
   Call Ivd.SetFieldValue("DOCUMENT_NO", "***ENDYEAR_" & uctlToDate.ShowDate)
   Call Ivd.SetFieldValue("DOCUMENT_DATE", uctlToDate.ShowDate)
   Call Ivd.SetFieldValue("DOCUMENT_TYPE", ADJUST_DOCTYPE)
   Call Ivd.SetFieldValue("COMMIT_FLAG", "Y")
   Call Ivd.SetFieldValue("EXCEPTION_FLAG", "N")
   Call Ivd.SetFieldValue("SALE_FLAG", "N")
   Call Ivd.SetFieldValue("ADJUST_FLAG", "N")
   Call Ivd.SetFieldValue("DOCUMENT_DESC", "ตั้งยอดปรับสิ้นปี " & uctlToDate.ShowDate)
   Call Ivd.SetFieldValue("DEPARTMENT_ID", -1)
   Call Ivd.SetFieldValue("CANCEL_FLAG", "N")
   
   For Each m_LotItem In InventoryBalBeforeAdjustColl
      I = I + 1
      txtPercent.Text = 80 + (MyDiff(I, InventoryBalBeforeAdjustColl.Count) * 15)
      prgProgress.Value = Val(txtPercent.Text)
      Me.Refresh
      DoEvents
      
      Set TempLotItem = New CLotItem
      TempLotItem.Flag = "A"
      TempLotItem.PART_ITEM_ID = m_LotItem.PART_ITEM_ID
      TempLotItem.LOCATION_ID = m_LotItem.LOCATION_ID
            
      Set TempLotItemSearch = GetObject("CLotItem", InventoryBalAferAdjustColl, Trim(m_LotItem.LOCATION_ID & "-" & m_LotItem.PART_ITEM_ID))
      If (m_LotItem.SUM_AMOUNT > TempLotItemSearch.SUM_AMOUNT) Then  'ก่อนปรับมากกว่าหลังปรับ ต้องปรับเพิ่ม
         TempLotItem.TX_AMOUNT = m_LotItem.SUM_AMOUNT - TempLotItemSearch.SUM_AMOUNT
         TempLotItem.MULTIPLIER = 1
         TempLotItem.TX_TYPE = "I"
      ElseIf (m_LotItem.SUM_AMOUNT < TempLotItemSearch.SUM_AMOUNT) Then
         TempLotItem.TX_AMOUNT = TempLotItemSearch.SUM_AMOUNT - m_LotItem.SUM_AMOUNT
         TempLotItem.MULTIPLIER = -1
         TempLotItem.TX_TYPE = "E"
      End If
      Set Pi = GetObject("CStockCode", TempPartColl, Trim(Str(m_LotItem.PART_ITEM_ID)))
      TempLotItem.UNIT_TRAN_ID = Pi.UNIT_CHANGE_ID
      TempLotItem.UNIT_MULTIPLE = 1

      Call Ivd.ImportExportItems.add(TempLotItem)
      
      Set TempLotItem = Nothing
   Next m_LotItem
   
   If Ivd.ImportExportItems.Count > 0 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, True, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Exit Function
      End If
   End If
   
   Set Ivd = Nothing
   Set TempPartColl = Nothing
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   Me.Refresh
   DoEvents
   Set m_LotItem = Nothing
   AdjustStockCode = True
   MasterInd = "1"
End Function
