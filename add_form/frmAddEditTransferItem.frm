VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTransferItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTransferItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   5325
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   9393
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextLookup uctlPartTypeLookup 
         Height          =   435
         Left            =   2235
         TabIndex        =   0
         Top             =   300
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtPrice 
         Height          =   435
         Left            =   2235
         TabIndex        =   5
         Top             =   2100
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   2235
         TabIndex        =   3
         Top             =   1650
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   2235
         TabIndex        =   1
         Top             =   750
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   2235
         TabIndex        =   2
         Top             =   1200
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   2250
         TabIndex        =   9
         Top             =   3960
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   2250
         TabIndex        =   6
         Top             =   2550
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlPartToLookup 
         Height          =   435
         Left            =   2235
         TabIndex        =   8
         Top             =   3480
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlPartTypeToLookup 
         Height          =   435
         Left            =   2235
         TabIndex        =   7
         Top             =   3000
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblPartTypeTo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   26
         Top             =   3030
         Width           =   1935
      End
      Begin VB.Label lblPartTo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   25
         Top             =   3540
         Width           =   1935
      End
      Begin Threed.SSCommand cmdUnit 
         Height          =   405
         Left            =   4320
         TabIndex        =   4
         Top             =   1650
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4965
         TabIndex        =   24
         Top             =   1740
         Width           =   2565
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   23
         Top             =   2610
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   4305
         TabIndex        =   22
         Top             =   2580
         Width           =   1005
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2055
         TabIndex        =   10
         Top             =   4500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   450
         TabIndex        =   21
         Top             =   4020
         Width           =   1695
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   4290
         TabIndex        =   20
         Top             =   2130
         Width           =   1005
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3720
         TabIndex        =   11
         Top             =   4500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTransferItem.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5370
         TabIndex        =   12
         Top             =   4500
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   435
         TabIndex        =   19
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   75
         TabIndex        =   18
         Top             =   330
         Width           =   2055
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   17
         Top             =   810
         Width           =   1935
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   435
         TabIndex        =   16
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   420
         TabIndex        =   15
         Top             =   1230
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAddEditTransferItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection


Private m_PartTypes As Collection
Private m_ToPartTypes As Collection
Private m_Parts As Collection
Private m_ToParts As Collection
Private m_Locations As Collection
Private m_ToLocations As Collection
Private m_Mr As CMasterRef
Public ParentForm As Object
Public DocumentType As INVENTORY_DOCTYPE

'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String

Private LotFlag As Boolean
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

   Call InitNormalLabel(lblPartType, MapText("จากประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblPart, MapText("จากวัตถุดิบ"))
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblPrice, MapText("ราคา"))
   Call InitNormalLabel(lblLocation, MapText("จากสถานที่จัดเก็บ"))
   Call InitNormalLabel(lblToLocation, MapText("ไปสถานที่จัดเก็บ"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblTotalPrice, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblUnit, MapText(""))
   
   Call InitNormalLabel(lblPartTypeTo, MapText("ไปยังประเภทวัตถุดิบ"))
   Call InitNormalLabel(lblPartTo, MapText("ไปยังวัตถุดิบ"))
   
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPrice.Enabled = False
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalPrice.Enabled = False
   
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
         Dim EnpAddr As CTransferItem

         Set EnpAddr = TempCollection.Item(ID)

         uctlPartTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeLookup.MyCombo, EnpAddr.ExportItem.PART_TYPE)
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, EnpAddr.ExportItem.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, EnpAddr.ExportItem.LOCATION_ID)
         
         uctlPartTypeToLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeToLookup.MyCombo, EnpAddr.ImportItem.PART_TYPE)
         uctlPartToLookup.MyCombo.ListIndex = IDToListIndex(uctlPartToLookup.MyCombo, EnpAddr.ImportItem.PART_ITEM_ID)
         uctlToLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlToLocationLookup.MyCombo, EnpAddr.ImportItem.LOCATION_ID)
         
         txtQuantity.Text = MyDiffEx(EnpAddr.ExportItem.TX_AMOUNT, EnpAddr.ExportItem.UNIT_MULTIPLE)
         txtPrice.Text = EnpAddr.ExportItem.AVG_PRICE * EnpAddr.ExportItem.UNIT_MULTIPLE
         txtTotalPrice.Text = EnpAddr.ExportItem.TOTAL_INCLUDE_PRICE
         UnitID = EnpAddr.ExportItem.UNIT_TRAN_ID
         Multiple = EnpAddr.ExportItem.UNIT_MULTIPLE
         UnitName = EnpAddr.ExportItem.UNIT_TRAN_NAME
         
         Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
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
      uctlPartTypeLookup.MyCombo.ListIndex = -1
      uctlPartLookup.MyCombo.ListIndex = -1
      txtQuantity.Text = ""
      txtPrice.Text = ""
      uctlLocationLookup.MyCombo.ListIndex = -1
      uctlToLocationLookup.MyCombo.ListIndex = -1
   End If
   
   Call QueryData(True)
   Call ParentForm.RefreshGrid(DocumentType, True)
   uctlPartTypeLookup.SetFocus
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
   If Not VerifyTextControl(lblPrice, txtPrice, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   
'   If uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) = _
'       uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex)) Then
'         glbErrorLog.LocalErrorMsg = "โรงเรือนเข้ากับโรงเรือนออกจะต้องแตกต่างกัน"
'         glbErrorLog.ShowUserError
'
'         Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If ShowMode = SHOW_ADD Then
'      If Not (LoadCheckBalance(Val(txtQuantity.Text) * Multiple, uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)), uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)), uctlPartLookup.MyTextBox.Text)) Then
'         SaveData = False
'         Exit Function
'      End If
   End If
   
   Dim EnpAddress As CTransferItem
   Dim Ei As CLotItem
   Dim II As CLotItem
   If ShowMode = SHOW_ADD Then
      Set Ei = New CLotItem
      Set II = New CLotItem
      Set EnpAddress = New CTransferItem
      
      Ei.Flag = "A"
      II.Flag = "A"
      EnpAddress.Flag = "A"
      
      Set EnpAddress.ExportItem = Ei
      Set EnpAddress.ImportItem = II

      Call TempCollection.add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
         EnpAddress.ExportItem.Flag = "E"
         EnpAddress.ImportItem.Flag = "E"
      End If
   End If
   
   
   If ShowMode = SHOW_EDIT Then
'      If Not (LoadCheckBalance(Val(txtQuantity.Text) * Multiple, uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)), uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)), uctlPartLookup.MyTextBox.Text, EnpAddress.ExportItem.GetFieldValue("LOT_ITEM_ID"))) Then
'         SaveData = False
'         Exit Function
'      End If
   End If
   
   EnpAddress.ExportItem.PART_TYPE = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   EnpAddress.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   
   EnpAddress.ExportItem.TX_AMOUNT = Val(txtQuantity.Text) * Multiple
   EnpAddress.ExportItem.AVG_PRICE = MyDiffEx(Val(txtPrice.Text), Multiple)
   EnpAddress.ExportItem.TOTAL_INCLUDE_PRICE = Val(txtTotalPrice.Text)
   
   EnpAddress.ExportItem.UNIT_TRAN_ID = UnitID
   EnpAddress.ExportItem.UNIT_MULTIPLE = Multiple
   EnpAddress.ExportItem.UNIT_TRAN_NAME = UnitName
   
   EnpAddress.ExportItem.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.ExportItem.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   EnpAddress.ExportItem.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.ExportItem.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.ExportItem.TX_TYPE = "E"
   EnpAddress.ExportItem.MULTIPLIER = -1
   
   EnpAddress.ImportItem.PART_TYPE = uctlPartTypeToLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeToLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.PART_ITEM_ID = uctlPartToLookup.MyCombo.ItemData(Minus2Zero(uctlPartToLookup.MyCombo.ListIndex))
   EnpAddress.ImportItem.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))
   
   EnpAddress.ImportItem.TX_AMOUNT = Val(txtQuantity.Text) * Multiple
   EnpAddress.ImportItem.AVG_PRICE = MyDiffEx(Val(txtPrice.Text), Multiple)
   EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = (txtQuantity.Text) * Val(txtPrice.Text)
   
   EnpAddress.ImportItem.PART_TYPE_NAME = uctlPartLookup.MyCombo.Text
   EnpAddress.ImportItem.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
   EnpAddress.ImportItem.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.ImportItem.PART_DESC = uctlPartLookup.MyCombo.Text
   EnpAddress.ImportItem.TX_TYPE = "I"
   EnpAddress.ImportItem.MULTIPLIER = 1
   
   EnpAddress.ImportItem.UNIT_TRAN_ID = UnitID
   EnpAddress.ImportItem.UNIT_MULTIPLE = Multiple
   EnpAddress.ImportItem.UNIT_TRAN_NAME = UnitName
   
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
   
   SendKeys ("{TAB}")

End Sub


Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadMaster(uctlPartTypeLookup.MyCombo, m_PartTypes, , , MASTER_STOCKTYPE)
      Set uctlPartTypeLookup.MyCollection = m_PartTypes
            
      Call LoadMaster(uctlPartTypeToLookup.MyCombo, m_ToPartTypes, , , MASTER_STOCKTYPE)
      Set uctlPartTypeToLookup.MyCollection = m_ToPartTypes
      
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call LoadMaster(uctlToLocationLookup.MyCombo, m_ToLocations, , , MASTER_LOCATION)
      Set uctlToLocationLookup.MyCollection = m_ToLocations
      
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
   Set m_ToLocations = New Collection
   Set m_Mr = New CMasterRef
   Set m_ToPartTypes = New Collection
   Set m_ToParts = New Collection
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PartTypes = Nothing
   Set m_Parts = Nothing
   Set m_Locations = Nothing
   Set m_ToLocations = Nothing
   Set m_Mr = Nothing
   Set m_ToPartTypes = Nothing
   Set m_ToParts = Nothing
End Sub
Private Sub txtQuantity_Change()
   m_HasModify = True
End Sub
Private Sub txtPrice_Change()
   m_HasModify = True
   txtTotalPrice.Text = Val(txtPrice.Text) * Val(txtQuantity.Text)
End Sub
Private Sub txtTotalPrice_Change()
   m_HasModify = True
End Sub
Private Sub uctlPartLookup_Change()
Dim PartItemID As Long
Dim Pi As CStockCode

   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      If ShowMode = SHOW_ADD Then
         uctlPartToLookup.MyCombo.ListIndex = IDToListIndex(uctlPartToLookup.MyCombo, PartItemID)
      End If
      
      Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(PartItemID)))
            
      UnitID = Pi.UNIT_ID
      Multiple = Pi.UNIT_AMOUNT
      UnitName = Pi.UNIT_NAME
      UnitMName = Pi.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   End If

   m_HasModify = True
End Sub

Private Sub uctlPartToLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartTypeLookup_Change()
Dim PartTypeID As Long

   PartTypeID = uctlPartTypeLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      If ShowMode = SHOW_ADD Then
         uctlPartTypeToLookup.MyCombo.ListIndex = IDToListIndex(uctlPartTypeToLookup.MyCombo, PartTypeID)
      End If
      
      Call LoadStockCode(uctlPartLookup.MyCombo, m_Parts, PartTypeID)
      Set uctlPartLookup.MyCollection = m_Parts
   End If
   
   m_HasModify = True
End Sub
Private Sub uctlPartTypeToLookup_Change()
Dim PartTypeID As Long

   PartTypeID = uctlPartTypeToLookup.MyCombo.ItemData(Minus2Zero(uctlPartTypeToLookup.MyCombo.ListIndex))
   
   If PartTypeID > 0 Then
      Call LoadStockCode(uctlPartToLookup.MyCombo, m_ToParts, PartTypeID)
      Set uctlPartToLookup.MyCollection = m_ToParts
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub
