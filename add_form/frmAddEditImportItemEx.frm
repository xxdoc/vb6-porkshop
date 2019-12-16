VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditImportItemEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditImportItemEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   7223
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtPrice 
         Height          =   435
         Left            =   1785
         TabIndex        =   5
         Top             =   2130
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1785
         TabIndex        =   3
         Top             =   1680
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1785
         TabIndex        =   2
         Top             =   1230
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtTotalPrice 
         Height          =   465
         Left            =   1800
         TabIndex        =   6
         Top             =   2580
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   820
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4485
         TabIndex        =   19
         Top             =   1740
         Width           =   2565
      End
      Begin Threed.SSCommand cmdUnit 
         Height          =   435
         Left            =   3840
         TabIndex        =   4
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditImportItemEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1838
         TabIndex        =   7
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditImportItemEx.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   18
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3870
         TabIndex        =   17
         Top             =   2610
         Width           =   465
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3488
         TabIndex        =   8
         Top             =   3300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditImportItemEx.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5138
         TabIndex        =   9
         Top             =   3300
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
         TabIndex        =   16
         Top             =   2190
         Width           =   1485
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   15
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   14
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   13
         Top             =   1740
         Width           =   1485
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   12
         Top             =   1320
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditImportItemEx"
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
Public TempCollection2 As Collection
Public COMMIT_FLAG As String
Public SupplierID As Long
Public DocumentType As INVENTORY_DOCTYPE
Public ParentForm As Object

Private m_PartTypes As Collection
Private m_Parts As Collection
Private m_Locations As Collection
Private m_Packagings As Collection
Private m_PartItemSpecs As Collection
Private m_PurchaseExpenses As Collection
Private m_Mr As CMasterRef

'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String
'------------------------------------------------------
Private LotFlag  As Boolean



'--------------------------------------------------
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
   If DocumentType = IMPORT_DOCTYPE Then
      Call InitNormalLabel(lblQuantity, MapText("ปริมาณนำเข้า"))
   ElseIf DocumentType = EXPORT_DOCTYPE Then
      Call InitNormalLabel(lblQuantity, MapText("ปริมาณที่เบิก"))
   End If
   Call InitNormalLabel(lblPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblLocation, MapText("สถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(lblUnit, MapText(""))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   If DocumentType = EXPORT_DOCTYPE Then
      txtTotalPrice.Enabled = False
   End If
   
   Call uctlPartLookup.MyTextBox.SetKeySearch("STOCK_NO")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUnit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdUnit, MapText("U"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CLotItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.LOCATION_ID)
         
         txtQuantity.Text = MyDiffEx(EnpAddr.TX_AMOUNT, EnpAddr.UNIT_MULTIPLE)
         txtPrice.Text = EnpAddr.AVG_PRICE * EnpAddr.UNIT_MULTIPLE
         txtTotalPrice.Text = EnpAddr.TOTAL_INCLUDE_PRICE
         UnitID = EnpAddr.UNIT_TRAN_ID
         Multiple = EnpAddr.UNIT_MULTIPLE
         UnitName = EnpAddr.UNIT_TRAN_NAME
         
         Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
         
         cmdOK.Enabled = (COMMIT_FLAG <> "Y")
         
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid(DocumentType, True)
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      'uctlPartTypeLookup.MyCombo.ListIndex = -1
      uctlPartLookup.MyCombo.ListIndex = -1
      txtQuantity.Text = ""
      txtPrice.Text = ""
      txtTotalPrice.Text = ""
      'uctlLocationLookup.MyCombo.ListIndex = -1
   End If
   
   Set TempCollection2 = New Collection
   
   Call QueryData(True)
   Call uctlPartTypeLookup.SetFocus
   Call ParentForm.RefreshGrid(DocumentType, True)
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
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, DocumentType = EXPORT_DOCTYPE) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblPrice, txtPrice, DocumentType = EXPORT_DOCTYPE) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If DocumentType = EXPORT_DOCTYPE Then
      If ShowMode = SHOW_ADD Then
'         If Not (LoadCheckBalance(Val(txtQuantity.Text) * Multiple, uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)), uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)), uctlPartLookup.MyTextBox.Text)) Then
'            Exit Function
'         End If
      End If
   End If
   
   Dim EnpAddress As CLotItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CLotItem
      EnpAddress.Flag = "A"
      Call TempCollection.add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If
   
   
   If DocumentType = EXPORT_DOCTYPE Then
      If ShowMode = SHOW_EDIT Then
'         If Not (LoadCheckBalance(Val(txtQuantity.Text) * Multiple, uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)), uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)), uctlPartLookup.MyTextBox.Text, EnpAddress.GetFieldValue("LOT_ITEM_ID"))) Then
'            Exit Function
'         End If
      End If
   End If
   
   EnpAddress.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   
   EnpAddress.TX_AMOUNT = Val(txtQuantity.Text) * Multiple
   EnpAddress.AVG_PRICE = MyDiffEx(Val(txtPrice.Text), Multiple)
   EnpAddress.TOTAL_INCLUDE_PRICE = Val(txtTotalPrice.Text)
   
   EnpAddress.UNIT_TRAN_ID = UnitID
   EnpAddress.UNIT_MULTIPLE = Multiple
   EnpAddress.UNIT_TRAN_NAME = UnitName
   
   EnpAddress.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   If DocumentType = IMPORT_DOCTYPE Then
      EnpAddress.TX_TYPE = "I"
      EnpAddress.MULTIPLIER = 1
   ElseIf DocumentType = EXPORT_DOCTYPE Then
      EnpAddress.TX_TYPE = "E"
      EnpAddress.MULTIPLIER = -1
   End If

   Set EnpAddress = Nothing
   SaveData = True
End Function
Private Sub cmdUnit_Click()
   frmChangeUnit.HeaderText = MapText("เปลี่ยนหน่วย")
   frmChangeUnit.UnitID = UnitID
   frmChangeUnit.Multiple = Multiple
   frmChangeUnit.UnitName = UnitName
   frmChangeUnit.UnitMName = UnitMName
   
   Load frmChangeUnit
   frmChangeUnit.Show 1
   
   UnitID = frmChangeUnit.UnitID
   Multiple = frmChangeUnit.Multiple
   UnitName = frmChangeUnit.UnitName
   UnitMName = frmChangeUnit.UnitMName
   
   Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   
   Unload frmChangeUnit
   Set frmChangeUnit = Nothing
   
   m_HasModify = True
   SendKeys ("{TAB}")
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadMaster(uctlPartTypeLookup.MyCombo, m_PartTypes, , , MASTER_STOCKTYPE)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
            
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
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
      'Call cmdAddLotItem_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdNext_Click
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
   Set m_Mr = New CMasterRef
   
   Set TempCollection2 = New Collection
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
   Set m_Mr = Nothing
   
   Set TempCollection2 = Nothing
End Sub

Private Sub txtPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   txtTotalPrice.Text = Val(txtQuantity.Text) * Val(txtPrice.Text)
End Sub

Private Sub txtTotalPrice_Change()
   m_HasModify = True
   txtPrice.Text = MyDiffEx(Val(txtTotalPrice.Text), Val(txtQuantity.Text))
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
Dim PartItemID As Long
Dim Pi As CStockCode
   If uctlPartLookup.MyCombo.ListCount <= 0 Then
      Exit Sub
   End If
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(PartItemID)))
      UnitID = Pi.UNIT_ID
      Multiple = Pi.UNIT_AMOUNT
      UnitName = Pi.UNIT_NAME
      UnitMName = Pi.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   End If

   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long
Static OldPartType  As Long
   
   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   If OldPartType <> PartTypeID Then
      OldPartType = PartTypeID
   Else
      Exit Sub
   End If
   
   If PartTypeID > 0 Then
      Call LoadStockCode(uctlPartLookup.MyCombo, m_Parts, PartTypeID)
      Set uctlPartLookup.MyCollection = m_Parts
   End If
   
   m_HasModify = True
   
End Sub
