VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditImportItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditImportItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7605
      Left            =   0
      TabIndex        =   31
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13414
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCalculateType 
         Height          =   510
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3900
         Width           =   2625
      End
      Begin prjWINPricing500.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   17
         Top             =   3900
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   5625
         TabIndex        =   16
         Top             =   3450
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1815
         TabIndex        =   21
         Top             =   4800
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   1800
         TabIndex        =   20
         Top             =   4350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextLookup uctlLayoutLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtEntryWeight 
         Height          =   435
         Left            =   1770
         TabIndex        =   4
         Top             =   1650
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtExitWeight 
         Height          =   435
         Left            =   5640
         TabIndex        =   5
         Top             =   1650
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtWeightAmount 
         Height          =   435
         Left            =   8760
         TabIndex        =   6
         Top             =   1620
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtPackageWeight 
         Height          =   435
         Left            =   5640
         TabIndex        =   8
         Top             =   2100
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtPercentHumid 
         Height          =   435
         Left            =   1770
         TabIndex        =   12
         Top             =   3000
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtHumid 
         Height          =   435
         Left            =   5640
         TabIndex        =   13
         Top             =   3000
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtOtherWeight 
         Height          =   435
         Left            =   8760
         TabIndex        =   9
         Top             =   2070
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtSupplierWeight 
         Height          =   435
         Left            =   8760
         TabIndex        =   14
         Top             =   2970
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtPackageAmount 
         Height          =   435
         Left            =   1770
         TabIndex        =   7
         Top             =   2100
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtPackageAmount1 
         Height          =   435
         Left            =   1770
         TabIndex        =   10
         Top             =   2550
         Width           =   1515
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextLookup uctlExpense1 
         Height          =   435
         Left            =   1830
         TabIndex        =   22
         Top             =   5250
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextLookup uctlExpense2 
         Height          =   435
         Left            =   1830
         TabIndex        =   24
         Top             =   5700
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtExpense1 
         Height          =   435
         Left            =   8460
         TabIndex        =   23
         Top             =   5220
         Width           =   1545
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtExpense2 
         Height          =   435
         Left            =   8460
         TabIndex        =   25
         Top             =   5670
         Width           =   1545
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtNetPrice 
         Height          =   435
         Left            =   1830
         TabIndex        =   26
         Top             =   6150
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtTotalWeight 
         Height          =   435
         Left            =   1770
         TabIndex        =   15
         Top             =   3450
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjWINPricing500.uctlTextBox txtExtraName 
         Height          =   435
         Left            =   5640
         TabIndex        =   11
         Top             =   2550
         Width           =   2565
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkBagReturn 
         Height          =   495
         Left            =   8760
         TabIndex        =   19
         Top             =   1110
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   873
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdPartItem 
         Height          =   405
         Left            =   7170
         TabIndex        =   1
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditImportItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblExtraName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4050
         TabIndex        =   61
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblTotalWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   60
         Top             =   3510
         Width           =   1485
      End
      Begin VB.Label Label10 
         Height          =   375
         Left            =   3900
         TabIndex        =   59
         Top             =   6120
         Width           =   1005
      End
      Begin VB.Label lblNetPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   6210
         Width           =   1485
      End
      Begin VB.Label Label8 
         Height          =   375
         Left            =   10110
         TabIndex        =   57
         Top             =   5670
         Width           =   435
      End
      Begin VB.Label lblExpenseAmount2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7290
         TabIndex        =   56
         Top             =   5730
         Width           =   1125
      End
      Begin VB.Label Label6 
         Height          =   375
         Left            =   10110
         TabIndex        =   55
         Top             =   5220
         Width           =   435
      End
      Begin VB.Label lblExpenseAmount1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7290
         TabIndex        =   54
         Top             =   5280
         Width           =   1125
      End
      Begin VB.Label lblExpense2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   53
         Top             =   5760
         Width           =   1485
      End
      Begin VB.Label lblExpense1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   52
         Top             =   5310
         Width           =   1485
      End
      Begin VB.Label lblPackageAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   51
         Top             =   2610
         Width           =   1605
      End
      Begin VB.Label lblActualPackageAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   50
         Top             =   2160
         Width           =   1605
      End
      Begin VB.Label lblCalculateType 
         Alignment       =   1  'Right Justify
         Caption         =   "lblFormulaType"
         Height          =   315
         Left            =   4380
         TabIndex        =   49
         Top             =   3990
         Width           =   1185
      End
      Begin VB.Label lblSupplierWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7170
         TabIndex        =   48
         Top             =   3030
         Width           =   1485
      End
      Begin VB.Label lblOtherWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7170
         TabIndex        =   47
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Label lblHumid 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4050
         TabIndex        =   46
         Top             =   3000
         Width           =   1485
      End
      Begin VB.Label lblPercentHumid 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   45
         Top             =   3060
         Width           =   1485
      End
      Begin VB.Label lblPackageWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3630
         TabIndex        =   44
         Top             =   2100
         Width           =   1875
      End
      Begin VB.Label lblWeightAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7170
         TabIndex        =   43
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lblExitWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4050
         TabIndex        =   42
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblEntryWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   41
         Top             =   1710
         Width           =   1485
      End
      Begin Threed.SSCommand cmdLayout 
         Height          =   405
         Left            =   7170
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditImportItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label lblLayout 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   40
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   39
         Top             =   4410
         Width           =   1485
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   5400
         TabIndex        =   38
         Top             =   2190
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3870
         TabIndex        =   37
         Top             =   4350
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3713
         TabIndex        =   28
         Top             =   6810
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditImportItem.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5363
         TabIndex        =   29
         Top             =   6810
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   36
         Top             =   3960
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   35
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   34
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4035
         TabIndex        =   33
         Top             =   3510
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   32
         Top             =   4860
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditImportItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public COMMIT_FLAG As String
Public SupplierID As Long

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Packagings As Collection
Private m_PartItemSpecs As Collection
Private m_PurchaseExpenses As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub CalculateTotalPrice(Ind As Long)
   If Ind = 1 Then
      txtTotalPrice.Text = Val(txtSupplierWeight.Text) * Val(txtPrice.Text)
   ElseIf Ind = 2 Then
      txtTotalPrice.Text = Val(txtWeightAmount.Text) * Val(txtPrice.Text)
   ElseIf Ind = 3 Then
      txtTotalPrice.Text = Val(txtTotalWeight.Text) * Val(txtPrice.Text)
   Else
      txtTotalPrice.Text = 0
   End If
End Sub

Private Sub cboCalculateType_Click()
Dim TempID As Long

   TempID = cboCalculateType.ItemData(Minus2Zero(cboCalculateType.ListIndex))
   Call CalculateTotalPrice(TempID)
   
   m_HasModify = True
End Sub

Private Sub cboCalculateType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkBagReturn_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblPartType, MapText("ประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblPart, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblQuantity, MapText("น้ำหนักนำเข้า"))
   Call InitNormalLabel(lblPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblLayout, MapText("ภาชนะบรรจุ"))
   Call InitNormalLabel(lblEntryWeight, MapText("น้ำหนักเข้า"))
   Call InitNormalLabel(lblExitWeight, MapText("น้ำหนักออก"))
   Call InitNormalLabel(lblWeightAmount, MapText("น้ำหนักรวม"))
   Call InitNormalLabel(lblActualPackageAmount, MapText("จำนวนรับจริง"))
   Call InitNormalLabel(lblPackageWeight, MapText("น.น. ภาชนะบรรจุ"))
   Call InitNormalLabel(lblOtherWeight, MapText("น้ำหนักอื่น ๆ"))
   Call InitNormalLabel(lblPercentHumid, MapText("% ความชื้น"))
   Call InitNormalLabel(lblHumid, MapText("น้ำหนักความชื้น"))
   Call InitNormalLabel(lblSupplierWeight, MapText("น้ำหนักผู้ขาย"))
   Call InitNormalLabel(lblCalculateType, MapText("คิดราคาแบบ"))
   Call InitNormalLabel(lblPackageAmount, MapText("จำนวนบรรจุ"))
   Call InitNormalLabel(lblExpense1, MapText("คชจ. จัดซื้อ 1"))
   Call InitNormalLabel(lblExpense2, MapText("คชจ. จัดซื้อ 2"))
   Call InitNormalLabel(lblExpenseAmount1, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblExpenseAmount2, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblNetPrice, MapText("ราคาสุทธิ"))
   Call InitNormalLabel(Label6, MapText("บาท"))
   Call InitNormalLabel(Label8, MapText("บาท"))
   Call InitNormalLabel(Label10, MapText("บาท"))
   Call InitNormalLabel(lblTotalWeight, MapText("น้ำหนักสุทธิ"))
   Call InitNormalLabel(lblExtraName, MapText("ชื่อน้ำหนักอื่น ๆ"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalPrice.Enabled = False
   Call txtEntryWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtExitWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtWeightAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPackageWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtOtherWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPercentHumid.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
'   txtPercentHumid.Enabled = False
   Call txtHumid.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtSupplierWeight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPackageAmount1.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtExpense1.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtExpense2.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtNetPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtNetPrice.Enabled = False
   Call txtTotalWeight.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtExtraName.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCheckBox(chkBagReturn, "คืนถุง")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdLayout.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPartItem.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboCalculateType)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdLayout, MapText("..."))
   Call InitMainButton(cmdPartItem, MapText("..."))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CLotItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.GetFieldValue("PART_TYPE"))
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.GetFieldValue("PART_ITEM_ID"))
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.GetFieldValue("LOCATION_ID"))
         
         txtEntryWeight.Text = EnpAddr.ENTRY_WEIGHT
         txtExitWeight.Text = EnpAddr.EXIT_WEIGHT
         txtWeightAmount.Text = EnpAddr.WEIGHT_AMOUNT
         txtPackageWeight.Text = EnpAddr.PACKAGE_WEIGHT
         txtOtherWeight.Text = EnpAddr.OTHER_WEIGHT
         txtPercentHumid.Text = EnpAddr.PERCENT_HUMID
         txtHumid.Text = EnpAddr.HUMID_WEIGHT
         txtPackageAmount1.Text = EnpAddr.PACKAGE_AMOUNT
         txtPackageAmount.Text = EnpAddr.ACTUAL_PKG_AMOUNT
   
         txtQuantity.Text = EnpAddr.TX_AMOUNT
         txtPrice.Text = EnpAddr.ACTUAL_UNIT_PRICE
         txtTotalPrice.Text = EnpAddr.TOTAL_ACTUAL_PRICE
         txtSupplierWeight.Text = EnpAddr.SUPPLIER_WEIGHT
         cboCalculateType.ListIndex = IDToListIndex(cboCalculateType, EnpAddr.CALCULATE_TYPE)
         uctlExpense1.MyCombo.ListIndex = IDToListIndex(uctlExpense1.MyCombo, EnpAddr.PUREXP_ID1)
         uctlExpense2.MyCombo.ListIndex = IDToListIndex(uctlExpense2.MyCombo, EnpAddr.PUREXP_ID2)
         txtExpense1.Text = EnpAddr.EXPENSE1
         txtExpense2.Text = EnpAddr.EXPENSE2
         txtTotalWeight.Text = EnpAddr.TOTAL_WEIGHT
         txtExtraName.Text = EnpAddr.EXTRA_NAME
         chkBagReturn.Value = FlagToCheck(EnpAddr.BAG_RETURN)
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdLayout_Click()
Dim OKClick As Boolean
Dim LayoutID As Long

   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   frmLayoutSearch.PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   frmLayoutSearch.LocationID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   Load frmLayoutSearch
   frmLayoutSearch.Show 1
   
   OKClick = frmLayoutSearch.OKClick
   LayoutID = frmLayoutSearch.LayoutID
   
   Unload frmLayoutSearch
   Set frmLayoutSearch = Nothing
   
   If OKClick Then
      uctlLayoutLookup.MyCombo.ListIndex = IDToListIndex(uctlLayoutLookup.MyCombo, LayoutID)
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyCombo(lblPartType, uctlPartTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPrice, txtPrice, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCalculateType, cboCalculateType, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CLotItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CLotItem
      EnpAddress.Flag = "A"
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If

   EnpAddress.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   EnpAddress.TX_AMOUNT = txtQuantity.Text
   EnpAddress.ACTUAL_UNIT_PRICE = txtPrice.Text
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.CALCULATE_FLAG = "Y"
   EnpAddress.TOTAL_ACTUAL_PRICE = Val(txtTotalPrice.Text)
   EnpAddress.TX_TYPE = "I"
   EnpAddress.PACKAGING_ID = uctlLayoutLookup.MyCombo.ItemData(Minus2Zero(uctlLayoutLookup.MyCombo.ListIndex))
   EnpAddress.ENTRY_WEIGHT = Val(txtEntryWeight.Text)
   EnpAddress.EXIT_WEIGHT = Val(txtExitWeight.Text)
   EnpAddress.WEIGHT_AMOUNT = Val(txtWeightAmount.Text)
   EnpAddress.OTHER_WEIGHT = Val(txtOtherWeight.Text)
   EnpAddress.PACKAGE_WEIGHT = Val(txtPackageWeight.Text)
   EnpAddress.PERCENT_HUMID = Val(txtPercentHumid.Text)
   EnpAddress.HUMID_WEIGHT = Val(txtHumid.Text)
   EnpAddress.SUPPLIER_WEIGHT = Val(txtSupplierWeight.Text)
   EnpAddress.CALCULATE_TYPE = cboCalculateType.ItemData(Minus2Zero(cboCalculateType.ListIndex))
   EnpAddress.PACKAGE_AMOUNT = Val(txtPackageAmount1.Text)
   EnpAddress.ACTUAL_PKG_AMOUNT = Val(txtPackageAmount.Text)
   EnpAddress.PUREXP_ID1 = uctlExpense1.MyCombo.ItemData(Minus2Zero(uctlExpense1.MyCombo.ListIndex))
   EnpAddress.PUREXP_ID2 = uctlExpense2.MyCombo.ItemData(Minus2Zero(uctlExpense2.MyCombo.ListIndex))
   EnpAddress.EXPENSE1 = Val(txtExpense1.Text)
   EnpAddress.EXPENSE2 = Val(txtExpense2.Text)
   EnpAddress.TOTAL_WEIGHT = Val(txtTotalWeight.Text)
   EnpAddress.EXTRA_NAME = txtExtraName.Text
   EnpAddress.BAG_RETURN = Check2Flag(chkBagReturn.Value)
   
   If EnpAddress.CALCULATE_TYPE = 1 Then
      EnpAddress.CALCULATE_WEIGHT = EnpAddress.SUPPLIER_WEIGHT
   ElseIf EnpAddress.CALCULATE_TYPE = 2 Then
      EnpAddress.CALCULATE_WEIGHT = EnpAddress.WEIGHT_AMOUNT
   ElseIf EnpAddress.CALCULATE_TYPE = 3 Then
      EnpAddress.CALCULATE_WEIGHT = EnpAddress.TOTAL_WEIGHT
   End If
   
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub cmdPartItem_Click()
Dim OKClick As Boolean
Dim TempCol As Collection
Dim Cs As CPartItem

   Set TempCol = New Collection
   
   Set frmQueryPartItem.TempCollection = TempCol
   frmQueryPartItem.ShowMode = SHOW_ADD
   Load frmQueryPartItem
   frmQueryPartItem.Show 1
   
   OKClick = frmQueryPartItem.OKClick
   
   Unload frmQueryPartItem
   Set frmQueryPartItem = Nothing
   
   If OKClick Then
      Set Cs = TempCol(1)
      uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, Cs.PART_TYPE)
      uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, Cs.PART_ITEM_ID)
      m_HasModify = True
   End If
   
   Set TempCol = Nothing
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadPackaging(uctlLayoutLookup.MyCombo, m_Packagings)
      Set uctlLayoutLookup.MyCollection = m_Packagings
      
      Call LoadPartType(uctlPartTypeLookup.MyCombo, m_PartTypes)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
      
      Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call LoadPurchaseExpense(uctlExpense1.MyCombo, m_PurchaseExpenses)
      Set uctlExpense1.MyCollection = m_PurchaseExpenses
      
      Call LoadPurchaseExpense(uctlExpense2.MyCombo, m_PurchaseExpenses)
      Set uctlExpense2.MyCollection = m_PurchaseExpenses
      
      Call InitCalculateType(cboCalculateType)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PartTypes = New Collection
   Set m_Parts = New Collection
   Set m_Locations = New Collection
   Set m_Packagings = New Collection
   Set m_PartItemSpecs = New Collection
   Set m_PurchaseExpenses = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_Packagings = Nothing
   Set m_PartItemSpecs = Nothing
   Set m_PurchaseExpenses = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtEntryWeight_Change()
   txtWeightAmount.Text = Val(txtEntryWeight.Text) - Val(txtExitWeight.Text)
   
   m_HasModify = True
End Sub

Private Sub txtExitWeight_Change()
   txtWeightAmount.Text = Val(txtEntryWeight.Text) - Val(txtExitWeight.Text)

   m_HasModify = True
End Sub

Private Sub txtExpense1_Change()
   m_HasModify = True
   txtNetPrice.Text = Val(txtTotalPrice.Text) + Val(txtExpense1.Text)
End Sub

Private Sub txtExpense1_GotFocus()
Dim TempID As Long
Dim Pg As CPurchaseExpense

   If Len(txtExpense1.Text) > 0 Then
      Exit Sub
   End If
   
   TempID = uctlExpense1.MyCombo.ItemData(Minus2Zero(uctlExpense1.MyCombo.ListIndex))
   If TempID > 0 Then
      Set Pg = GetPurchaseExpense(m_PurchaseExpenses, Trim(Str(TempID)))
      If Val(Pg.PUREXP_NO) = 1 Then 'พวกภาษี
         txtExpense1.Text = Pg.EXPENSE_RATE * Val(txtTotalPrice.Text)
      Else
         txtExpense1.Text = Pg.EXPENSE_RATE * Val(txtWeightAmount.Text)
      End If
   End If
End Sub

Private Sub txtExpense2_Change()
   m_HasModify = True
   txtNetPrice.Text = Val(txtTotalPrice.Text) + Val(txtExpense2.Text)
End Sub

Private Sub txtExpense2_GotFocus()
Dim TempID As Long
Dim Pg As CPurchaseExpense

   If Len(txtExpense2.Text) > 0 Then
      Exit Sub
   End If

   TempID = uctlExpense2.MyCombo.ItemData(Minus2Zero(uctlExpense2.MyCombo.ListIndex))
   If TempID > 0 Then
      Set Pg = GetPurchaseExpense(m_PurchaseExpenses, Trim(Str(TempID)))
      If Val(Pg.PUREXP_NO) = 1 Then 'พวกภาษี
         txtExpense2.Text = Pg.EXPENSE_RATE * Val(txtTotalPrice.Text)
      Else
         txtExpense2.Text = Pg.EXPENSE_RATE * Val(txtWeightAmount.Text)
      End If
   End If
End Sub

Private Sub txtExtraName_Change()
   m_HasModify = True
End Sub

Private Sub txtHumid_Change()
   m_HasModify = True
   txtTotalWeight.Text = Val(txtWeightAmount.Text) - Flag2Money(Check2Flag(chkBagReturn.Value), Val(txtPackageWeight.Text)) - Val(txtOtherWeight.Text) - Val(txtHumid.Text)
End Sub

Private Sub txtHumid_GotFocus()
Dim Ps As CPartItemSpec
Dim HumidPercent As Double
Dim HumidRate As Double
Dim TempAmount As Double

   If Len(txtHumid.Text) > 0 Then
      Exit Sub
   End If
   
   HumidRate = 0
   HumidPercent = Val(txtPercentHumid.Text)
   For Each Ps In m_PartItemSpecs
      If (Ps.FROM_RATE <= HumidPercent) And _
           (Ps.TO_RATE >= HumidPercent) Then
           
            HumidRate = Ps.HUMIDITY_WEIGHT
            Exit For
         End If
   Next Ps
   
   TempAmount = Val(txtWeightAmount.Text)
   txtHumid.Text = HumidRate * TempAmount / 1000
End Sub

Private Sub txtNetPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtOtherWeight_Change()
   txtTotalWeight.Text = Val(txtWeightAmount.Text) - Flag2Money(Check2Flag(chkBagReturn.Value), Val(txtPackageWeight.Text)) - Val(txtOtherWeight.Text) - Val(txtHumid.Text)
   m_HasModify = True
End Sub

Private Sub txtPackageAmount1_Change()
   m_HasModify = True
End Sub

Private Sub txtPackageWeight_Change()
   m_HasModify = True
   txtTotalWeight.Text = Val(txtWeightAmount.Text) - Flag2Money(Check2Flag(chkBagReturn.Value), Val(txtPackageWeight.Text)) - Val(txtOtherWeight.Text) - Val(txtHumid.Text)
End Sub

Private Sub txtPackageWeight_GotFocus()
Dim TempID As Long
Dim Pg As CPackaging

   If Len(txtPackageWeight.Text) > 0 Then
      Exit Sub
   End If
   
   TempID = uctlLayoutLookup.MyCombo.ItemData(Minus2Zero(uctlLayoutLookup.MyCombo.ListIndex))
   If TempID > 0 Then
      Set Pg = GetPackaging(m_Packagings, Trim(Str(TempID)))
      txtPackageWeight.Text = Pg.WEIGHT_RATE * Val(txtPackageAmount.Text)
   End If
End Sub

Private Sub txtPercentHumid_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalWeight_Change()
   m_HasModify = True
   txtTotalPrice.Text = Val(txtPrice.Text) * Val(txtTotalWeight.Text)
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtPrice_Change()
Dim TempID As Long

   m_HasModify = True
   
   If cboCalculateType.ListIndex > 0 Then
      TempID = cboCalculateType.ItemData(Minus2Zero(cboCalculateType.ListIndex))
      Call CalculateTotalPrice(TempID)
   End If
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub txtSupplierWeight_Change()
   m_HasModify = True
End Sub

Private Sub txtTotalPrice_Change()
   m_HasModify = True
   txtNetPrice.Text = Val(txtTotalPrice.Text) + Val(txtExpense1.Text) + Val(txtExpense2.Text)
End Sub

Private Function Flag2Money(Flag As String, Value As Double) As Double
   If Flag = "Y" Then
      Flag2Money = 0
   Else
      Flag2Money = Value
   End If
End Function

Private Sub txtWeightAmount_Change()
   m_HasModify = True
   txtTotalWeight.Text = Val(txtWeightAmount.Text) - Flag2Money(Check2Flag(chkBagReturn.Value), Val(txtPackageWeight.Text)) - Val(txtOtherWeight.Text) - Val(txtHumid.Text)
End Sub

Private Sub uctlExpense1_Change()
   m_HasModify = True
End Sub

Private Sub uctlExpense2_Change()
   m_HasModify = True
End Sub

Private Sub uctlLayoutLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub CalculateItemAmount(Ind As Long)
Dim TempWeight As Double
Dim HumidWeight As Double

   If (Ind = 1) Or (Ind = 0) Then
      txtWeightAmount.Text = Val(txtEntryWeight.Text) - Val(txtExitWeight.Text)
   End If

   TempWeight = Val(txtWeightAmount.Text) - Val(txtPackageWeight.Text) - Val(txtOtherWeight.Text)
   HumidWeight = TempWeight * Val(txtPercentHumid.Text) / 100

   If (Ind = 2) Or (Ind = 0) Then
      txtQuantity.Text = TempWeight - HumidWeight
      txtHumid.Text = HumidWeight
   End If

   If (Ind = 3) Or (Ind = 0) Then
      txtHumid.Text = HumidWeight
   End If

   If (Ind = 4) Or (Ind = 0) Then
      txtQuantity.Text = Val(txtWeightAmount.Text) - Val(txtHumid.Text)
   End If
End Sub

Private Sub uctlPartLookup_Change()
Dim Sp As CSupplierSpec
Dim PartItemID As Long

   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Call LoadPartItemSpec(Nothing, m_PartItemSpecs, PartItemID)
   End If
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Dim Pt As CPartType

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Set Pt = GetPartType(m_PartTypes, Trim(Str(PartTypeID)))
      Call LoadPartItem(uctlPartLookup.MyCombo, m_Parts, PartTypeID, "N")
      Set uctlPartLookup.MyCollection = m_Parts
   
         Call LoadLocation(uctlLocationLookup.MyCombo, m_Locations, 2, , , Pt.PART_GROUP_ID)
         Set uctlLocationLookup.MyCollection = m_Locations
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox2_Change()
   m_HasModify = True
End Sub
