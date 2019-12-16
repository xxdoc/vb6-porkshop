VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmBillingDoc1 
   BackColor       =   &H80000000&
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
   Icon            =   "frmBillingDoc1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   13860
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10575
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   18653
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2670
         Width           =   2985
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2190
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6555
         Left            =   60
         TabIndex        =   7
         Top             =   3210
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   11562
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmBillingDoc1.frx":27A2
         Column(2)       =   "frmBillingDoc1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBillingDoc1.frx":290E
         FormatStyle(2)  =   "frmBillingDoc1.frx":2A6A
         FormatStyle(3)  =   "frmBillingDoc1.frx":2B1A
         FormatStyle(4)  =   "frmBillingDoc1.frx":2BCE
         FormatStyle(5)  =   "frmBillingDoc1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBillingDoc1.frx":2D5E
      End
      Begin PorkShop.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   840
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1290
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1740
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlFromDate 
         Height          =   405
         Left            =   6900
         TabIndex        =   20
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   6900
         TabIndex        =   22
         Top             =   1260
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtReceiptDoNo 
         Height          =   435
         Left            =   6900
         TabIndex        =   25
         Top             =   2160
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   767
      End
      Begin VB.Label lblReceiptDoNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   26
         Top             =   2220
         Width           =   1905
      End
      Begin Threed.SSCheck ChkCancelFlag 
         Height          =   435
         Left            =   6960
         TabIndex        =   24
         Top             =   2640
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   23
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4920
         TabIndex        =   21
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   2730
         Width           =   1755
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   15
         Top             =   2250
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9750
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9750
         TabIndex        =   6
         Top             =   2490
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3300
         TabIndex        =   10
         Top             =   9870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   30
         TabIndex        =   8
         Top             =   9870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1650
         TabIndex        =   9
         Top             =   9870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12135
         TabIndex        =   12
         Top             =   9870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10485
         TabIndex        =   11
         Top             =   9870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBillingDoc1.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmBillingDoc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_BillingDoc As CBillingDoc
Private m_TempBillingDoc As CBillingDoc
Private m_Rs As ADODB.Recordset

Public OKClick As Boolean
Public DocumentType As SELL_BILLING_DOCTYPE
Public ReceiptType As Long
Public Area As Long
Public HeaderText As String

Private CurrentIndex As Long
Private RightClickCollection As Collection
Private Sub ChkCancelFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim TempStr As String
   
   If Area = 1 Then
      TempStr = "(ขาย)"
   ElseIf Area = 2 Then
      TempStr = "(ซื้อ)"
   End If
   
   If DocumentType = INVOICE_DOCTYPE Then
      Set oMenu = New CPopupMenu
      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SellSubMenuItems)
      If lMenuChosen > 0 Then
         frmAddEditBillingDoc.DocumentSubType = lMenuChosen
      Else
         Exit Sub
      End If
   End If
   If DocumentType = PO_DOCTYPE Then
      Set oMenu = New CPopupMenu
      lMenuChosen = oMenu.AddMenu(glbGuiConfigs.PoSubMenuItems)
      If lMenuChosen > 0 Then
         frmAddEditBillingDoc.DocumentSubType = lMenuChosen
      Else
         Exit Sub
      End If
   End If
   If DocumentType = RECEIPT3_DOCTYPE Then
      frmAddEditBillingDocRcp.Area = Area
      frmAddEditBillingDocRcp.DocumentType = DocumentType
      frmAddEditBillingDocRcp.HeaderText = MapText("เพิ่มข้อมูล" & SellDoctype2Text(DocumentType))
      frmAddEditBillingDocRcp.ShowMode = SHOW_ADD
      Load frmAddEditBillingDocRcp
      frmAddEditBillingDocRcp.Show 1
      
      OKClick = frmAddEditBillingDocRcp.OKClick
   
      Unload frmAddEditBillingDocRcp
      Set frmAddEditBillingDocRcp = Nothing
   Else
      frmAddEditBillingDoc.Area = Area
      frmAddEditBillingDoc.DocumentType = DocumentType
      frmAddEditBillingDoc.HeaderText = MapText("เพิ่มข้อมูล" & SellDoctype2Text(DocumentType))
      frmAddEditBillingDoc.ShowMode = SHOW_ADD
      Load frmAddEditBillingDoc
      frmAddEditBillingDoc.Show 1
      
      OKClick = frmAddEditBillingDoc.OKClick
   
      Unload frmAddEditBillingDoc
      Set frmAddEditBillingDoc = Nothing
   End If
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtDocumentNo.Text = ""
   txtCustomerCode.Text = ""
   txtPartNo.Text = ""
   
   uctlFromDate.ShowDate = -1
   uctlToDate.ShowDate = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   CurrentIndex = GridEX1.Row
   
   If Not VerifyLockDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = QUOATATION_DOCTYPE Or DocumentType = RETURN_DOCTYPE _
   Or DocumentType = S_PO_DOCTYPE Or DocumentType = S_INVOICE_DOCTYPE Or DocumentType = S_RECEIPT1_DOCTYPE Or DocumentType = S_QUOATATION_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      If Not VerifyLockInvoiceDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   Else
      If Not VerifyLockReceiptDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If

   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_BillingDoc.BILLING_DOC_ID = ID
   If Not glbDaily.DeleteBillingDoc(m_BillingDoc, IsOK, True, glbErrorLog) Then
      m_BillingDoc.BILLING_DOC_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
   
   CurrentIndex = GridEX1.Row - 1
   Call QueryData(True)
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim TempStr As String
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   CurrentIndex = GridEX1.Row
   
   If Area = 1 Then
      TempStr = "(ขาย)"
   ElseIf Area = 2 Then
      TempStr = "(ซื้อ)"
   End If
      
   If DocumentType = RECEIPT3_DOCTYPE Then
      frmAddEditBillingDoc.Area = Area
      frmAddEditBillingDocRcp.ID = ID
      frmAddEditBillingDocRcp.DocumentType = DocumentType
      frmAddEditBillingDocRcp.HeaderText = MapText("แก้ไขข้อมูล" & SellDoctype2Text(DocumentType))
      frmAddEditBillingDocRcp.ShowMode = SHOW_EDIT
      Load frmAddEditBillingDocRcp
      frmAddEditBillingDocRcp.Show 1
      
      OKClick = frmAddEditBillingDocRcp.OKClick
   
      Unload frmAddEditBillingDocRcp
      Set frmAddEditBillingDocRcp = Nothing
   Else
      frmAddEditBillingDoc.Area = Area
      frmAddEditBillingDoc.ID = ID
      frmAddEditBillingDoc.DocumentType = DocumentType
      frmAddEditBillingDoc.HeaderText = MapText("แก้ไขข้อมูล" & SellDoctype2Text(DocumentType))
      frmAddEditBillingDoc.ShowMode = SHOW_EDIT
      Load frmAddEditBillingDoc
      frmAddEditBillingDoc.Show 1
      
      OKClick = frmAddEditBillingDoc.OKClick
      
      Unload frmAddEditBillingDoc
      Set frmAddEditBillingDoc = Nothing
   End If
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call GenerateRightClick(RightClickCollection)
      
      Call InitBillingDocOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      uctlFromDate.ShowDate = Now
      uctlToDate.ShowDate = Now
      
      Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_BillingDoc.BILLING_DOC_ID = -1
      m_BillingDoc.DOCUMENT_NO = PatchWildCard(txtDocumentNo.Text)
      m_BillingDoc.APAR_CODE = PatchWildCard(txtCustomerCode.Text)
      m_BillingDoc.FROM_DATE = uctlFromDate.ShowDate
      m_BillingDoc.TO_DATE = uctlToDate.ShowDate
      m_BillingDoc.DOCUMENT_TYPE = DocumentType
      m_BillingDoc.DOC_ID_NO = txtReceiptDoNo.Text
      If DocumentType <> RECEIPT3_DOCTYPE Then
         m_BillingDoc.APAR_IND = Area    'เป็นตัวแยกว่าซื้อหรือขาย"
      End If
      m_BillingDoc.STOCK_NO = PatchWildCard(txtPartNo.Text)
      m_BillingDoc.ORDER_BY = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_BillingDoc.ORDER_TYPE = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      
      m_BillingDoc.CANCEL_FLAG = Check2Flag(ChkCancelFlag.Value)
      
      If cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex)) <= 0 Then
         m_BillingDoc.ORDER_TYPE = 2
      End If
      If Not glbDaily.QueryBillingDoc(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
   
   If CurrentIndex > 0 Then
      Call GridEX1.SetFocus
      Call GridEX1.MoveToRowIndex(CurrentIndex)
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 1600
   Col.Caption = MapText("เลขที่เอกสาร")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("วันที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "CANCEL"
   
   If Area = 1 And DocumentType <> RECEIPT3_DOCTYPE Then
      Set Col = GridEX1.Columns.add '4
      Col.Width = 1200
      Col.Caption = MapText("รหัสลูกค้า")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 4600
      Col.Caption = MapText("ชื่อลูกค้า")
   ElseIf Area = 2 And DocumentType <> RECEIPT3_DOCTYPE Then
      Set Col = GridEX1.Columns.add '4
      Col.Width = 1200
      Col.Caption = MapText("รหัสซัพพลายเออร์")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 4600
      Col.Caption = MapText("ชื่อซัพพลายเออร์")
   End If
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1450
   Col.TextAlignment = jgexAlignRight
   If DocumentType = RECEIPT2_DOCTYPE Then
      Col.Caption = MapText("หักหนี้")
   Else
      Col.Caption = MapText("จำนวนเงิน")
   End If
   
   If DocumentType = INVOICE_DOCTYPE Or DocumentType = PO_DOCTYPE Then
      Set Col = GridEX1.Columns.add '9
      Col.Width = 3000
      Col.Caption = MapText("ประเภทใบส่งสินค้าย่อย")
   End If
   
   If DocumentType = RECEIPT3_DOCTYPE Then
      Set Col = GridEX1.Columns.add '9
      Col.Width = 3000
      Col.Caption = MapText("รายละเอียด")
   End If
   
Set Col = Nothing
Set fmsTemp = Nothing
GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   
   Call InitGrid
   
   If Area = 1 Then
      Call InitNormalLabel(lblCustomerCode, MapText("รหัสลูกค้า"))
   Else
      Call InitNormalLabel(lblCustomerCode, MapText("รหัสผู้ค้า"))
   End If
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสสินค้า"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblReceiptDoNo, MapText("เลขที่ใบส่งของชำระ"))
   
   If DocumentType = RECEIPT3_DOCTYPE Then
      txtCustomerCode.Enabled = False
      lblCustomerCode.Enabled = False
   End If
   Call InitCheckBox(ChkCancelFlag, "CANCEL")
   
   Call txtPartNo.SetKeySearch("STOCK_NO")
   Call txtCustomerCode.SetKeySearch("CUSTOMER_CODE")
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   m_HasActivate = False
   
   
   Set m_BillingDoc = New CBillingDoc
   Set m_TempBillingDoc = New CBillingDoc
   Set m_Rs = New ADODB.Recordset
   
   Set RightClickCollection = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_BillingDoc = Nothing
   Set m_TempBillingDoc = Nothing
   Set RightClickCollection = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As CPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim TempID2 As Long
Dim BD As CBillingDoc
Dim IsOK As Boolean
Dim OKClick As Boolean

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
   TempID1 = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New CPopupMenu
      
      lMenuChosen = oMenu.AddMenu(RightClickCollection)
      
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      Set oMenu = New CPopupMenu
      If DocumentType = INVOICE_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.SellSubMenuItems)
      ElseIf DocumentType = PO_DOCTYPE Then
         lMenuChosen = oMenu.AddMenu(glbGuiConfigs.PoSubMenuItems)
      End If
      If lMenuChosen > 0 Then
         Call EnableForm(Me, False)
         Call glbDaily.StartTransaction
         Set BD = New CBillingDoc
         BD.BILLING_DOC_ID = TempID1
         If lMenuChosen = 1 Then       'PO
            BD.DOCUMENT_SUB_TYPE = -1
         Else
            BD.DOCUMENT_SUB_TYPE = lMenuChosen
         End If
         Call BD.UpdateDocSubType
         Call glbDaily.CommitTransaction
      End If
      Set oMenu = Nothing
   End If
      
   Set BD = Nothing
   Call QueryData(True)
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(4)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I  As Long
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempBillingDoc.PopulateFromRS(1, m_Rs)
   
   I = 0
   I = I + 1
   Values(I) = m_TempBillingDoc.BILLING_DOC_ID
   I = I + 1
   Values(I) = m_TempBillingDoc.DOCUMENT_NO & "(" & m_TempBillingDoc.BILLING_DOC_PACK & ")"
   I = I + 1
   Values(I) = DateToStringExtEx2(m_TempBillingDoc.DOCUMENT_DATE)
   I = I + 1
   Values(I) = m_TempBillingDoc.CANCEL_FLAG
   
   If DocumentType <> RECEIPT3_DOCTYPE Then
      I = I + 1
      Values(I) = m_TempBillingDoc.APAR_CODE
      I = I + 1
      Values(I) = m_TempBillingDoc.APAR_NAME
   End If
   I = I + 1
   If DocumentType = RECEIPT2_DOCTYPE Or DocumentType = BILLS_DOCTYPE Or DocumentType = RECEIPT3_DOCTYPE Or DocumentType = S_RECEIPT2_DOCTYPE Or DocumentType = S_BILLS_DOCTYPE Then
      Values(I) = FormatNumber(m_TempBillingDoc.PAID_AMOUNT - m_TempBillingDoc.CREDIT_AMOUNT + m_TempBillingDoc.DEBIT_AMOUNT - m_TempBillingDoc.SUBTRACT_AMOUNT + m_TempBillingDoc.ADDITION_AMOUNT)
   ElseIf DocumentType = CN_DOCTYPE Or DocumentType = DN_DOCTYPE Or DocumentType = S_CN_DOCTYPE Or DocumentType = S_DN_DOCTYPE Then
      Values(I) = FormatNumber(m_TempBillingDoc.PAY_AMOUNT + m_TempBillingDoc.VAT_AMOUNT)
   Else
      Values(I) = FormatNumber(m_TempBillingDoc.TOTAL_PRICE - m_TempBillingDoc.DISCOUNT_AMOUNT - m_TempBillingDoc.EXT_DISCOUNT_AMOUNT + m_TempBillingDoc.VAT_AMOUNT)
   End If
   
   If DocumentType = INVOICE_DOCTYPE Then
      I = I + 1
      Values(I) = m_TempBillingDoc.DOCUMENT_SUB_TYPE_NAME
   ElseIf DocumentType = PO_DOCTYPE Then
      I = I + 1
      If m_TempBillingDoc.DOCUMENT_SUB_TYPE <= 0 Then '  PO
         Values(I) = "ขายสด"
      Else
         Values(I) = m_TempBillingDoc.DOCUMENT_SUB_TYPE_NAME
      End If
   End If
   
   If DocumentType = RECEIPT3_DOCTYPE Then
      I = I + 1
      Values(I) = m_TempBillingDoc.NOTE
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GenerateRightClick(Col As Collection)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '===
   If (DocumentType = PO_DOCTYPE Or DocumentType = INVOICE_DOCTYPE) Then
      Set D = New CMenuItem
      D.KEY_ID = 1
      D.KEYWORD = "ปรับประเภทใบส่งสินค้า "
      Call Col.add(D)
      Set D = Nothing
      
      Set D = New CMenuItem
      D.KEY_ID = 2
      D.KEYWORD = "-"
      Call Col.add(D)
      Set D = Nothing
   End If
   
   If (DocumentType = INVOICE_DOCTYPE Or DocumentType = RETURN_DOCTYPE) Then
      Set D = New CMenuItem
      D.KEY_ID = 3
      D.KEYWORD = "ปรับประเภทเอกสารนำกลับ/(PN) "
      Call Col.add(D)
      Set D = Nothing
      
      Set D = New CMenuItem
      D.KEY_ID = 4
      D.KEYWORD = "-"
      Call Col.add(D)
      Set D = Nothing
   End If
   
   TempCount = 0
   For Each D In Col
      TempCount = TempCount + 1
      If TempCount = Col.Count Then
         If D.KEYWORD = "-" Then
            Col.Remove (TempCount)
         End If
      End If
   Next D
   '====
End Sub

