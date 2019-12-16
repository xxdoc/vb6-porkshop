VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditEnpAddress 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditEnpAddress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6165
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   10874
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCountry 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3870
         Width           =   2685
      End
      Begin PorkShop.uctlTextBox txtSoi 
         Height          =   435
         Left            =   1710
         TabIndex        =   1
         Top             =   720
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtHomeNo 
         Height          =   435
         Left            =   1710
         TabIndex        =   0
         Top             =   270
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtVillage 
         Height          =   435
         Left            =   1710
         TabIndex        =   3
         Top             =   1620
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtMoo 
         Height          =   435
         Left            =   1710
         TabIndex        =   2
         Top             =   1170
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtRoad 
         Height          =   435
         Left            =   1710
         TabIndex        =   4
         Top             =   2070
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtAmphur 
         Height          =   435
         Left            =   1710
         TabIndex        =   6
         Top             =   2970
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtDistrict 
         Height          =   435
         Left            =   1710
         TabIndex        =   5
         Top             =   2520
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtProvince 
         Height          =   435
         Left            =   1710
         TabIndex        =   7
         Top             =   3420
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtFax 
         Height          =   435
         Left            =   1710
         TabIndex        =   11
         Top             =   4770
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtPhone 
         Height          =   435
         Left            =   1710
         TabIndex        =   10
         Top             =   4320
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtZipcode 
         Height          =   435
         Left            =   6900
         TabIndex        =   8
         Top             =   3390
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3750
         TabIndex        =   13
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditEnpAddress.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5400
         TabIndex        =   14
         Top             =   5400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkBangkok 
         Height          =   405
         Left            =   7680
         TabIndex        =   12
         Top             =   4800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblZipcode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5460
         TabIndex        =   28
         Top             =   3450
         Width           =   1335
      End
      Begin VB.Label lblPhone 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4350
         Width           =   1485
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   4830
         Width           =   1485
      End
      Begin VB.Label lblVillage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label lblCountry 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   120
         TabIndex        =   24
         Top             =   3900
         Width           =   1485
      End
      Begin VB.Label lblHomeNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label lblSoi 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblMoo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1230
         Width           =   1485
      End
      Begin VB.Label lblRoad 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Label lblAmphur 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3030
         Width           =   1485
      End
      Begin VB.Label lblDistrict 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2580
         Width           =   1485
      End
      Begin VB.Label lblProvince 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditEnpAddress"
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

Private m_MasterRef As CMasterRef

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
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
      
   Call InitNormalLabel(lblHomeNo, MapText("บ้านเลขที่"))
   Call InitNormalLabel(lblSoi, MapText("ซอย"))
   Call InitNormalLabel(lblMoo, MapText("หมู่"))
   Call InitNormalLabel(lblVillage, MapText("หมู่บ้าน"))
   Call InitNormalLabel(lblRoad, MapText("ถนน"))
   Call InitNormalLabel(lblDistrict, MapText("แขวง/ตำบล"))
   Call InitNormalLabel(lblAmphur, MapText("เขต/อำเภอ"))
   Call InitNormalLabel(lblProvince, MapText("จังหวัด"))
   Call InitNormalLabel(lblCountry, MapText("ประเทศ"))
   Call InitNormalLabel(lblPhone, MapText("โทรศัพท์"))
   Call InitNormalLabel(lblFax, MapText("แฟกซ์"))
   Call InitNormalLabel(lblZipcode, MapText("รหัสไปรษณีย์"))
   
   Call txtHomeNo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtSoi.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtMoo.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtVillage.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtRoad.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDistrict.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtAmphur.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtProvince.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPhone.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtFax.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtZipcode.SetTextLenType(TEXT_STRING, glbSetting.ZIP_TYPE)
   
   Call InitCheckBox(chkBangkok, "เมืองหลวง")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboCountry)
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Addr As CAddress
         Dim EnpAddr As CEnterpriseAddress
         Set EnpAddr = TempCollection.Item(ID)
         Set Addr = EnpAddr.Addresses(1)
         
         txtHomeNo.Text = Addr.GetFieldValue("HOME")
         txtSoi.Text = Addr.GetFieldValue("SOI")
         txtMoo.Text = Addr.GetFieldValue("MOO")
         txtVillage.Text = Addr.GetFieldValue("VILLAGE")
         txtRoad.Text = Addr.GetFieldValue("ROAD")
         txtDistrict.Text = Addr.GetFieldValue("DISTRICT")
         txtAmphur.Text = Addr.GetFieldValue("AMPHUR")
         txtProvince.Text = Addr.GetFieldValue("PROVINCE")
         txtZipcode.Text = Addr.GetFieldValue("ZIPCODE")
         txtPhone.Text = Addr.GetFieldValue("PHONE1")
         txtFax.Text = Addr.GetFieldValue("FAX1")
         cboCountry.ListIndex = IDToListIndex(cboCountry, Addr.GetFieldValue("COUNTRY_ID"))
         chkBangkok.Value = FlagToCheck(Addr.GetFieldValue("BANGKOK_FLAG"))
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblHomeNo, txtHomeNo, False) Then
      Exit Function
   End If
'   If Not VerifyCombo(lblCountry, cboCountry, False) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Addr As CAddress
   Dim EnpAddress As CEnterpriseAddress
   If ShowMode = SHOW_ADD Then
      Set Addr = New CAddress
      Set EnpAddress = New CEnterpriseAddress
      Call EnpAddress.Addresses.add(Addr)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      Set Addr = EnpAddress.Addresses.Item(1)
   End If
   
   Call Addr.SetFieldValue("HOME", txtHomeNo.Text)
   Call Addr.SetFieldValue("SOI", txtSoi.Text)
   Call Addr.SetFieldValue("MOO", txtMoo.Text)
   Call Addr.SetFieldValue("VILLAGE", txtVillage.Text)
   Call Addr.SetFieldValue("ROAD", txtRoad.Text)
   Call Addr.SetFieldValue("DISTRICT", txtDistrict.Text)
   Call Addr.SetFieldValue("AMPHUR", txtAmphur.Text)
   Call Addr.SetFieldValue("PROVINCE", txtProvince.Text)
   Call Addr.SetFieldValue("ZIPCODE", txtZipcode.Text)
   Call Addr.SetFieldValue("COUNTRY_ID", cboCountry.ItemData(Minus2Zero(cboCountry.ListIndex)))
   Call Addr.SetFieldValue("COUNTRY_NAME", cboCountry.Text)
   Call Addr.SetFieldValue("FAX1", txtFax.Text)
   Call Addr.SetFieldValue("PHONE1", txtPhone.Text)
   Call Addr.SetFieldValue("BANGKOK_FLAG", Check2Flag(chkBangkok.Value))
   
   If ShowMode = SHOW_ADD Then
      Addr.Flag = "A"
      EnpAddress.Flag = "A"
      Call TempCollection.add(EnpAddress)
   Else
      If Addr.Flag <> "A" Then
         Addr.Flag = "E"
      End If
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If
   
   Set Addr = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboCountry, , , , MASTER_COUNTRY)
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
   Set m_MasterRef = New CMasterRef
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_MasterRef = Nothing
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

Private Sub txtMoo_Change()
   m_HasModify = True
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

Private Sub txtVillage_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub
