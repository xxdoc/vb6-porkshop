VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWinPricingMain 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "JasmineUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWinPricingMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":24B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":2D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3666
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":3980
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":425A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":4B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":540E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWinPricingMain.frx":5CE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel SSPanel1 
         Height          =   795
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1402
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSFrame LogoFrame 
            Height          =   855
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   1508
            _Version        =   131073
         End
         Begin VB.Label lblDateTime 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9360
            TabIndex        =   14
            Top             =   0
            Width           =   2505
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   555
            Left            =   9660
            TabIndex        =   13
            Top             =   6390
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   979
            _Version        =   131073
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmWinPricingMain.frx":5FF7
            Caption         =   "SSCommand1"
            ButtonStyle     =   3
         End
         Begin VB.Label lblCompany 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblCompany"
            BeginProperty Font 
               Name            =   "JasmineUPC"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   5295
            TabIndex        =   12
            Top             =   120
            Width           =   1755
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   7755
         Left            =   0
         TabIndex        =   2
         Top             =   750
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   13679
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   840
            Top             =   0
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   0
            Top             =   0
         End
         Begin MSComctlLib.TreeView trvMain 
            Height          =   4035
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   7117
            _Version        =   393217
            Indentation     =   882
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "JasmineUPC"
               Size            =   14.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblUserName 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            TabIndex        =   8
            Top             =   5580
            Width           =   3045
         End
         Begin VB.Label lblUserGroup 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            TabIndex        =   7
            Top             =   6090
            Width           =   3045
         End
         Begin VB.Label lblVersion 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   360
            TabIndex        =   6
            Top             =   6600
            Width           =   3045
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   465
            Left            =   1920
            TabIndex        =   5
            Top             =   7170
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   820
            _Version        =   131073
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdPasswd 
            Height          =   465
            Left            =   330
            TabIndex        =   4
            Top             =   7170
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   820
            _Version        =   131073
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   795
         Left            =   3450
         TabIndex        =   9
         Top             =   750
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   1402
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSFrame fraGeneric 
         Height          =   1455
         Left            =   4800
         TabIndex        =   10
         Top             =   1860
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2566
         _Version        =   131073
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand cmdGeneric 
            Height          =   885
            Index           =   0
            Left            =   720
            TabIndex        =   1
            Top             =   300
            Visible         =   0   'False
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   1561
            _Version        =   131073
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "SSCommand2"
         End
      End
   End
End
Attribute VB_Name = "frmWinPricingMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
 
Private m_Sp As CSystemParam
Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset

Public HeaderText As String
Private m_MustAsk As Boolean

Private m_Journals As Collection
Private TempCollection  As Collection
Private m_JobProcessMenus As Collection
Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String
Dim I As Long
   
   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT_EX
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False

   I = 0
   
   Set Node = trvMain.Nodes.add(, tvwFirst, ROOT_TREE, MapText("ระบบงานทั้งหมด"), 1)
   Node.Expanded = True
   Node.Selected = True

   '==
   I = I + 1
   Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", I & "." & MapText("ระบบข้อมูลผู้ใช้งาน"), 4, 4)
   Node.Expanded = False
   '==
   
   I = I + 1
   Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", I & "." & MapText("ระบบข้อมูลหลัก"), 2, 2)
   Node.Expanded = False

   I = I + 1
   Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", I & "." & MapText("ระบบข้อมูลส่วนกลาง"), 6, 6)
   Node.Expanded = False
   
   If glbGuiConfigs.VerifyGuiConfig("LEDGER_VIEW") Then
      I = I + 1
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", I & "." & MapText("ระบบบริหารบัญชี การเงิน"), 8, 8)
      Node.Expanded = False
   End If
   
   If glbGuiConfigs.VerifyGuiConfig("INVENTORY_VIEW") Then
      I = I + 1
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", I & "." & MapText("ระบบบริหารคลัง"), 3, 3)
      Node.Expanded = False
   End If
   
   If glbGuiConfigs.VerifyGuiConfig("PACKAGE_VIEW") Then
      I = I + 1
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-12", I & "." & MapText("ระบบการตั้งราคาสินค้า"), 5, 5)
      Node.Expanded = False
   End If
   
   If glbGuiConfigs.VerifyGuiConfig("PRODUCTION_VIEW") Then
      I = I + 1
      Set Node = trvMain.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-9", I & "." & MapText("ระบบการผลิต"), 10, 10)
      Node.Expanded = False
   End If
End Sub
Private Sub InitFormLayout()
   Call InitNormalLabel(lblUsername, MapText("ผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("เวอร์ชัน : ") & glbParameterObj.Version & " (FireBird) ", RGB(0, 0, 255))
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.BackStyle = 1
   lblDateTime.BackColor = RGB(255, 255, 255)
   
   lblCompany.Caption = MapText(glbEnterPrise.ENTERPRISE_NAME & "  " & glbEnterPrise.SHORT_NAME)
   'Me.Picture = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   LogoFrame.PictureBackground = LoadPicture(glbParameterObj.CompanyLogo)
   
   'LogoFrame.Visible = True
   LogoFrame.Visible = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Me.Caption = glbGuiConfigs.ShowWindowCaption()
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdPasswd, MapText("โปรแกรม"))
   
   lblCompany.ForeColor = RGB(0, 0, 255)
   lblCompany.BackColor = RGB(255, 255, 255)
   
   Call InitMainTreeview
End Sub
Private Sub cmdExit_Click()
   Unload Me
End Sub
Private Sub GenerateMasterMenuItem(Col As Collection, MasterArea As MASTER_TYPE)
Dim Mr As CMasterRef
Dim MI As CMenuItem
Dim TempCol As Collection
Dim I As Long

   Set Col = Nothing
   Set Col = New Collection
   
   Set TempCol = New Collection
   Set Mr = New CMasterRef
   
   Call LoadMaster(Nothing, TempCol, , , MasterArea)
   
   I = 0
   For Each Mr In TempCol
      I = I + 1
      Set MI = New CMenuItem
      MI.KEY_ID = Mr.KEY_ID
      MI.KEYWORD = Mr.KEY_NAME
      Call Col.add(MI)
      Set MI = Nothing
   
      If I < TempCol.Count Then
         Set MI = New CMenuItem
         MI.KEY_ID = -1
         MI.KEYWORD = "-"
         Call Col.add(MI)
         Set MI = Nothing
      End If
   Next Mr
   
   Set Mr = Nothing
   Set TempCol = Nothing
End Sub

Private Sub GenerateInventoryDocItems(Col As Collection)
Dim MI As CMenuItem
Dim TempCol As Collection
Dim I As INVENTORY_DOCTYPE
Dim j As Long

   Set Col = Nothing
   Set Col = New Collection
      
   I = 0
   j = 0
   For I = IMPORT_DOCTYPE To ADJUST_DOCTYPE
      j = j + 1
      Set MI = New CMenuItem
      MI.KEY_ID = I
      MI.KEYWORD = Doctype2Text(I)
      Call Col.add(MI)
      Set MI = Nothing
      
      If j < ADJUST_DOCTYPE Then
         Set MI = New CMenuItem
         MI.KEY_ID = -1
         MI.KEYWORD = "-"
         Call Col.add(MI)
         Set MI = Nothing
      End If
   Next I
   
'   Set MI = New CMenuItem
'   MI.KEY_ID = -1
'   MI.KEYWORD = "-"
'   Call Col.add(MI)
'   Set MI = Nothing
'
'   Set MI = New CMenuItem
'   MI.KEY_ID = 1000
'   MI.KEYWORD = "เอกสารการผลิต"
'   Call Col.add(MI)
   
   ' ที่ไม่เอาไปใส่ที่ InventoryDoc เนื่องจากว่ามันจะไปมีผลการทบต่อ การปรับปรืมาณและราคาเฉลี่ย โดยให้ ID มากกว่า 100
   Set MI = New CMenuItem
   MI.KEY_ID = -1
   MI.KEYWORD = "-"
   Call Col.add(MI)
   Set MI = Nothing
   
   Set MI = New CMenuItem
   MI.KEY_ID = 100
   MI.KEYWORD = "เอกสารการตรวจนับยอดสต็อค"
   Call Col.add(MI)
   Set MI = Nothing
   
   Set TempCol = Nothing
End Sub
Private Sub GenerateInventoryBarcode(Col As Collection)
Dim MI As CMenuItem
Dim TempCol As Collection
Dim I As INVENTORY_DOCTYPE
Dim j As Long

   Set Col = Nothing
   Set Col = New Collection
   
   Set MI = New CMenuItem
   MI.KEY_ID = IMPORT_DOCTYPE
   MI.KEYWORD = "(1)ใบรับเข้าคลัง/โรงงาน/CN"
   Call Col.add(MI, Trim(Str(MI.KEY_ID)))
   Set MI = Nothing
      
   Set MI = New CMenuItem
   MI.KEY_ID = EXPORT_DOCTYPE
   MI.KEYWORD = "(2) เบิก /ใบตัด Stock ฝากขาย สำรอง/ใบตัด Stock ชิม ส่งเสริม สูญเสีย "
   Call Col.add(MI, Trim(Str(MI.KEY_ID)))
   Set MI = Nothing
   
   Set MI = New CMenuItem
   MI.KEY_ID = ADJUST_DOCTYPE
   MI.KEYWORD = "(3) ผลิต"
   Call Col.add(MI, Trim(Str(MI.KEY_ID)))
   Set MI = Nothing
   
   Set MI = New CMenuItem
   MI.KEY_ID = TRANSFER_DOCTYPE
   MI.KEYWORD = "(4)ใบโอนจากโรงงาน/ฝากขาย/ใบเก็บสำรองไว้ที่ Sale/ใบรับคืนจากฝากขาย/รับคืนจากสำรอง ฝากขาย"
   Call Col.add(MI, Trim(Str(MI.KEY_ID)))
   Set MI = Nothing
   
   Set TempCol = Nothing
End Sub
Private Sub GenerateSellBillingDocItems(Col As Collection)
Dim MI As CMenuItem
Dim TempCol As Collection
Dim I As SELL_BILLING_DOCTYPE
Dim j As Long

   Set Col = Nothing
   Set Col = New Collection
      
   I = 0
   j = 0
   For I = PO_DOCTYPE To RECEIPT3_DOCTYPE
      j = j + 1
      Set MI = New CMenuItem
      MI.KEY_ID = I
      MI.KEYWORD = SellDoctype2Text(I)
      Call Col.add(MI)
      Set MI = Nothing
   
      If j < (RECEIPT3_DOCTYPE - 1) Then
         Set MI = New CMenuItem
         MI.KEY_ID = -1
         MI.KEYWORD = "-"
         Call Col.add(MI)
         Set MI = Nothing
      End If
   Next I
         
   Set TempCol = Nothing
End Sub
Private Sub GenerateBuyBillingDocItems(Col As Collection)
Dim MI As CMenuItem
Dim TempCol As Collection
Dim I As SELL_BILLING_DOCTYPE
Dim j As Long

   Set Col = Nothing
   Set Col = New Collection
      
   I = 0
   j = 0
   For I = S_PO_DOCTYPE To S_BILLS_DOCTYPE
      j = j + 1
      Set MI = New CMenuItem
      MI.KEY_ID = I
      MI.KEYWORD = SellDoctype2Text(I)
      Call Col.add(MI)
      Set MI = Nothing
   
      If j < (BILLS_DOCTYPE - 1) Then
         Set MI = New CMenuItem
         MI.KEY_ID = -1
         MI.KEYWORD = "-"
         Call Col.add(MI)
         Set MI = Nothing
      End If
   Next I

   Set TempCol = Nothing
End Sub
Private Sub cmdGeneric_Click(Index As Integer)
Dim Key As String
Dim Caption As String
Dim MenuSelected As Long
Dim MenuSelected2 As Long
Dim oMenu As CPopupMenu
Dim lNewmenu As Long
Dim DocumentType As Long

   Set oMenu = New CPopupMenu
   
   Key = cmdGeneric(Index).Tag
   Caption = cmdGeneric(Index).Caption
   
   If Key = "ADMIN_GROUP" Then
      If Not VerifyAccessRight("ADMIN_GROUP") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Load frmUserGroup
      frmUserGroup.Show 1
      
      Unload frmUserGroup
      Set frmUserGroup = Nothing
   ElseIf Key = "ADMIN_USER" Then
      If Not VerifyAccessRight("ADMIN_USER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Load frmUser
      frmUser.Show 1
      
      Unload frmUser
      Set frmUser = Nothing
   ElseIf Key = "MAIN_ENTERPRISE" Then
      If Not VerifyAccessRight("MAIN_ENTERPRISE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmAddEditEnterprise.HeaderText = MapText("ข้อมูลบริษัท")
      Load frmAddEditEnterprise
      frmAddEditEnterprise.Show 1
      
      Unload frmAddEditEnterprise
      Set frmAddEditEnterprise = Nothing
   ElseIf Key = "MAIN_CUSTOMER" Then
      If Not VerifyAccessRight("MAIN_CUSTOMER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmApArMas.ApArInd = 1
      frmApArMas.HeaderText = MapText("ข้อมูลลูกค้า")
      Load frmApArMas
      frmApArMas.Show 1
      
      Unload frmApArMas
      Set frmApArMas = Nothing
   ElseIf Key = "MAIN_SUPPLIER" Then
      If Not VerifyAccessRight("MAIN_SUPPLIER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmApArMas.ApArInd = 2
      frmApArMas.HeaderText = MapText("ข้อมูลผู้ค้า")
      Load frmApArMas
      frmApArMas.Show 1
      
      Unload frmApArMas
      Set frmApArMas = Nothing
   ElseIf Key = "MAIN_EMPLOYEE" Then
      If Not VerifyAccessRight("MAIN_EMPLOYEE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Load frmEmployee
      frmEmployee.Show 1
      
      Unload frmEmployee
      Set frmEmployee = Nothing
   ElseIf Key = "MAIN_REPORT" Then
      If Not VerifyAccessRight("MAIN_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 3
      Load frmSummaryReport
      frmSummaryReport.Show 1
   
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   ElseIf Key = "MASTER_MAIN" Then
      If Not VerifyAccessRight("MASTER_MAIN") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      frmMasterMain.HeaderText = MapText("ข้อมูลหลักส่วนกลาง")
      frmMasterMain.MasterMode = 1
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_INVENTORY" Then
      If Not VerifyAccessRight("MASTER_INVENTORY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmMasterMain.HeaderText = MapText("ข้อมูลหลักระบบคลัง")
      frmMasterMain.MasterMode = 3
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_LEDGER" Then
      If Not VerifyAccessRight("MASTER_LEDGER") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmMasterMain.HeaderText = MapText("ข้อมูลหลักระบบบัญชี")
      frmMasterMain.MasterMode = 4
      Load frmMasterMain
      frmMasterMain.Show 1
      
      Unload frmMasterMain
      Set frmMasterMain = Nothing
   ElseIf Key = "MASTER_REPORT" Then
      If Not VerifyAccessRight("MASTER_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 2
      Load frmSummaryReport
      frmSummaryReport.Show 1
   
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   ElseIf Key = "INVENTORY_PART" Then
      If Not VerifyAccessRight("INVENTORY_PART") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call GenerateMasterMenuItem(m_Journals, MASTER_STOCKGROUP)
      MenuSelected = oMenu.AddMenu(m_Journals)
      If MenuSelected <= 0 Then
         Exit Sub
      End If
      
      frmPartItem.PartGroupID = MenuSelected
      frmPartItem.HeaderText = MapText("ข้อมูลรหัสสต็อค")
      Load frmPartItem
      frmPartItem.Show 1
      
      Unload frmPartItem
      Set frmPartItem = Nothing
   ElseIf Key = "INVENTORY_DOC" Then
      If Not VerifyAccessRight("INVENTORY_DOC") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call GenerateInventoryDocItems(m_Journals)
      MenuSelected = oMenu.AddMenu(m_Journals)
      If MenuSelected <= 0 Then
         Exit Sub
      End If
      
      If MenuSelected = 100 Then
         If Not VerifyAccessRight("INVENTORY_DOC_เอกสารการตรวจนับยอดสต็อค") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         frmBalanceVerify.HeaderText = "เอกสารการตรวจนับยอดสต็อค"
         Load frmBalanceVerify
         frmBalanceVerify.Show 1
         
         Unload frmBalanceVerify
         Set frmBalanceVerify = Nothing
      Else
         If Not VerifyAccessRight("INVENTORY_DOC_" & Doctype2Text(MenuSelected)) Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         Call SeparateInventorySubTypeColl(TempCollection, MenuSelected)
         If TempCollection.Count > 0 Then
            Set oMenu = Nothing
            Set oMenu = New CPopupMenu
            MenuSelected2 = oMenu.AddMenu(TempCollection)
            If MenuSelected2 = 0 Then
               Exit Sub
            End If
         End If
         frmInventoryDoc.InventorySubType = MenuSelected2
         frmInventoryDoc.DocumentType = MenuSelected
         frmInventoryDoc.HeaderText = Doctype2Text(MenuSelected)
         Load frmInventoryDoc
         frmInventoryDoc.Show 1
         
         Unload frmInventoryDoc
         Set frmInventoryDoc = Nothing
      End If
    
   ElseIf Key = "INVENTORY_REPORT" Then
      If Not VerifyAccessRight("INVENTORY_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
         
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 6
      Load frmSummaryReport
      frmSummaryReport.Show 1
   
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
      
   ElseIf Key = "LEDGER_SELL" Then
      If Not VerifyAccessRight("LEDGER_SELL") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Call GenerateSellBillingDocItems(m_Journals)
      MenuSelected = oMenu.AddMenu(m_Journals)
      If MenuSelected <= 0 Then
         Exit Sub
      End If
      
      If Not VerifyAccessRight("LEDGER_SELL_" & SellDoctype2Text(MenuSelected)) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
            
      frmBillingDoc1.Area = 1
      frmBillingDoc1.DocumentType = MenuSelected
      frmBillingDoc1.HeaderText = SellDoctype2Text(MenuSelected)
      Load frmBillingDoc1
      frmBillingDoc1.Show 1
      
      Unload frmBillingDoc1
      Set frmBillingDoc1 = Nothing
   ElseIf Key = "LEDGER_BUY" Then
      If Not VerifyAccessRight("LEDGER_BUY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call GenerateBuyBillingDocItems(m_Journals)
      MenuSelected = oMenu.AddMenu(m_Journals)
      If MenuSelected <= 0 Then
         Exit Sub
      End If
      
      If Not VerifyAccessRight("LEDGER_BUY_" & SellDoctype2Text(MenuSelected)) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmBillingDoc1.Area = 2
      frmBillingDoc1.DocumentType = MenuSelected
      frmBillingDoc1.HeaderText = SellDoctype2Text(MenuSelected)
      Load frmBillingDoc1
      frmBillingDoc1.Show 1

      Unload frmBillingDoc1
      Set frmBillingDoc1 = Nothing
   ElseIf Key = "LEDGER_CASH" Then
      If Not VerifyAccessRight("LEDGER_CASH") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      MenuSelected = oMenu.AddMenu(glbGuiConfigs.CashMenuItems)
      If MenuSelected <= 0 Then
         Exit Sub
      End If
      
      If Not VerifyAccessRight("LEDGER_CASH_" & CashDocTypeToText(MenuSelected)) Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If MenuSelected = CHEQUE_REV Then
         frmCheque.Area = 1
         frmCheque.HeaderText = CashDocTypeToText(CHEQUE_REV)
         Load frmCheque
         frmCheque.Show 1
   
         Unload frmCheque
         Set frmCheque = Nothing
      ElseIf MenuSelected = CHEQUE_PAY Then
         frmCheque.Area = 2
         frmCheque.HeaderText = CashDocTypeToText(CHEQUE_PAY)
         Load frmCheque
         frmCheque.Show 1
   
         Unload frmCheque
         Set frmCheque = Nothing
      ElseIf MenuSelected = CASH_DEPOSIT Then
         frmCashDoc.DocumentType = MenuSelected
         frmCashDoc.HeaderText = CashDocTypeToText(MenuSelected)
         Load frmCashDoc
         frmCashDoc.Show 1
         
         Unload frmCashDoc
         Set frmCashDoc = Nothing
         Exit Sub
      ElseIf MenuSelected = POST_CHEQUE Then
         frmCashDoc.DocumentType = MenuSelected
         frmCashDoc.HeaderText = CashDocTypeToText(MenuSelected)
         Load frmCashDoc
         frmCashDoc.Show 1
         
         Unload frmCashDoc
         Set frmCashDoc = Nothing
         Exit Sub
      End If
      
   ElseIf Key = "LEDGER_REPORT" Then
      If Not VerifyAccessRight("LEDGER_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 5
      Load frmSummaryReport
      frmSummaryReport.Show 1
      
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   ElseIf Key = "LEDGER_PROGRAM" Then
      If Not VerifyAccessRight("LEDGER_PROGRAM") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      Call GenerateBillingProgram(m_Journals)
      MenuSelected = oMenu.AddMenu(m_Journals)
      If MenuSelected <= 0 Then
         Exit Sub
      End If
      
      If MenuSelected = 1 Then
         If Not VerifyAccessRight("LEDGER_PROGRAM_คัดลอกของมูลจากเอกสารใบสั่งซื้อมาเป็นบิลขาย ตามวันที่ส่งของ") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
         Load frmCopyPoToDo
         frmCopyPoToDo.Show 1
         
         Unload frmCopyPoToDo
         Set frmCopyPoToDo = Nothing
      End If
   
   
      
   ElseIf Key = "PACKAGE_DATA" Then
      If Not VerifyAccessRight("PACKAGE_DATA") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      frmPackage.HeaderText = "ข้อมูลการตั้งราคาขายสินค้า"
      Load frmPackage
      frmPackage.Show 1

      Unload frmPackage
      Set frmPackage = Nothing
   
   ElseIf Key = "COST_STOCK-AMOUNT" Then
      If Not VerifyAccessRight("COST_STOCK-AMOUNT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Debug.Print
      
      frmAdjustStockAmount.HeaderText = "ปรับยอด STOCK เป็นชุด"

      Load frmAdjustStockAmount
      frmAdjustStockAmount.Show 1

      Unload frmAdjustStockAmount
      Set frmAdjustStockAmount = Nothing
   ElseIf Key = "PRODUCT_FORMULA" Then
      If Not VerifyAccessRight("PRODUCT_FORMULA") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmFormula.HeaderText = Caption
      Load frmFormula
      frmFormula.Show 1
   
      Unload frmFormula
      Set frmFormula = Nothing
   ElseIf Key = "PRODUCT_JOB" Then
      If Not VerifyAccessRight("PRODUCT_JOB") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmJob1.HeaderText = Caption
      Load frmJob1
      frmJob1.Show 1
   
      Unload frmJob1
      Set frmJob1 = Nothing
   ElseIf Key = "PRODUCT_REPORT" Then
      If Not VerifyAccessRight("PRODUCT_REPORT") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      frmSummaryReport.HeaderText = Caption
      frmSummaryReport.MasterMode = 4
      Load frmSummaryReport
      frmSummaryReport.Show 1
   
      Unload frmSummaryReport
      Set frmSummaryReport = Nothing
   End If
   
   Set oMenu = Nothing
End Sub
Private Sub cmdPasswd_Click()
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim OKClick As Boolean
Dim iCount As Long
   
   
   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("เปลี่ยนรหัสผ่าน", "-", "คอนฟิคเลขที่เอกสาร", "-", "กำหนดวันที่เอกสาร")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
     If Not VerifyAccessRight("PROGRAM_โปรแกรมเปลี่ยนรหัสผ่าน") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      
      Load frmChangePassword
      frmChangePassword.Show 1

      Unload frmChangePassword
      Set frmChangePassword = Nothing
   ElseIf lMenuChosen = 3 Then
      If Not VerifyAccessRight("PROGRAM_ตั้งเลขที่เอกสาร") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      frmConfigDoc.HeaderText = "ตั้งเลขที่เอกสาร"
      Load frmConfigDoc
      frmConfigDoc.Show 1
      
      Unload frmConfigDoc
      Set frmConfigDoc = Nothing
   ElseIf lMenuChosen = 5 Then
      If Not VerifyAccessRight("PROGRAM_กำหนดวันที่เอกสาร") Then
         Call EnableForm(Me, True)
         Exit Sub
     End If
      frmLockDate.HeaderText = "กำหนดวันที่เอกสาร"
      Load frmLockDate
      frmLockDate.Show 1
      
      Unload frmLockDate
      Set frmLockDate = Nothing
   End If
End Sub
Private Sub Form_Activate()
Dim OKClick As Boolean
Dim iCount As Long
Dim Package As CPackageDetail
Dim CUS As CAPARMas
Dim Mr  As CMasterRef
Dim TempEmp As CEmployee
   
   Set Package = New CPackageDetail
   Set CUS = New CAPARMas
   Set Mr = New CMasterRef
   Set TempEmp = New CEmployee
   
   If Not m_HasActivate Then
      m_HasActivate = True
      Call PatchDB
      Load frmLogin
      frmLogin.Show 1
      
      OKClick = frmLogin.OKClick

      Unload frmLogin
      Set frmLogin = Nothing
            
      'Call Shell("C:\WINDOWS\system32\calc.exe ", vbMaximizedFocus)
      
      glbEnterPrise.ENTERPRISE_ID = -1
      Call glbEnterPrise.QueryData(1, m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call glbEnterPrise.PopulateFromRS(1, m_Rs)
         lblCompany.Caption = MapText(glbEnterPrise.ENTERPRISE_NAME & "  " & glbEnterPrise.SHORT_NAME)
      End If
      
      glbLockDate.LOCK_DATE_ID = -1
      glbLockDate.LOCK_TYPE = 1
      Call glbLockDate.QueryData(1, m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call glbLockDate.PopulateFromRS(1, m_Rs)
      End If
      
      CUS.APAR_IND = 1
      Call LoadApArMas(CUS, Nothing, m_CustomerColl)
      Set CUS = Nothing
      
      Set CUS = New CAPARMas
      CUS.APAR_IND = 2
      Call LoadApArMas(CUS, Nothing, m_SupplierColl)
      
      Call LoadEmployee(TempEmp, Nothing, m_EmployeeColl)
      
      Call LoadPackageDetail(Package, Nothing, LoadPackageColl)
      
      Call LoadMaster(Nothing, m_LocationColl, , , MASTER_LOCATION)
      
      Call LoadMaster(Nothing, InventorySubTypecoll, , , MASTER_INVENTORY_SUB_TYPE)
      
      Set CUS = Nothing
      Set Package = Nothing
      Set Mr = Nothing
      Set TempEmp = Nothing
      
      If Not OKClick Then
         m_MustAsk = False
         Unload Me
      Else
         Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
      End If
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
   
   Call InitFormLayout
   Set m_Journals = New Collection
   Set m_JobProcessMenus = New Collection
   Set TempCollection = New Collection
   Set m_Rs = New ADODB.Recordset
   
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านต้องการออกจากโปรแกรมใช่หรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
      Else
         Cancel = True
      End If
   End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
      
   SSPanel1.Width = ScaleWidth
   SSFrame2.Width = ScaleWidth
   SSFrame2.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth - SSFrame1.Width
   lblDateTime.Left = ScaleWidth - lblDateTime.Width - 50
   lblCompany.Left = (ScaleWidth - lblCompany.Width) / 2
   SSFrame1.Top = SSPanel1.Height
   pnlHeader.Top = SSPanel1.Height
   SSFrame1.Height = ScaleHeight - SSPanel1.Height
   
    cmdExit.Top = SSFrame1.Height - cmdExit.Height - 50
    cmdPasswd.Top = cmdExit.Top
    lblUsername.Top = SSFrame1.Height - 2200
    lblUserGroup.Top = SSFrame1.Height - 1600
    lblVersion.Top = SSFrame1.Height - 1000
   
    fraGeneric.Width = pnlHeader.Width * 2 / 3
    fraGeneric.Left = ((pnlHeader.Width - fraGeneric.Width) / 2) + SSFrame1.Width
    cmdGeneric(0).Width = fraGeneric.Width * 5 / 6
    cmdGeneric(0).Left = (fraGeneric.Width - cmdGeneric(0).Width) / 2
    
    trvMain.Height = SSFrame1.Height - 3300
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Call ReleaseAll
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Journals = Nothing
   Set m_JobProcessMenus = Nothing
   Set m_Sp = Nothing
   Set TempCollection = Nothing
End Sub
Private Sub Timer1_Timer()
   Timer1.Enabled = False
   
   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   lblUsername.Caption = MapText("ผู้ใช้ : ") & " " & glbUser.USER_NAME
   lblUserGroup.Caption = MapText("กลุ่มผู้ใช้ : ") & " " & glbUser.GROUP_NAME

   Timer1.Enabled = True   ' เปิด - ปิด timer ก่อน build program
   
End Sub

Private Sub Timer2_Timer()
   Call SetMousePos(100, 244)
   Call LeftClick
   'Sleep 2000
   Call SetMousePos(1028, 197)
   Call LeftClick
   'Sleep 2000
'   Call SetMousePos(1250, 250)
'   Call LeftClick
   'Debug.Print GetX & "-" & GetY
End Sub

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node Is Nothing Then
      Exit Sub
   End If
   
   pnlHeader.Caption = Node.Text
   If Node.Key = ROOT_TREE & " 1-0" Then
      Call InitCommandLayout(glbGuiConfigs.AdminCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-1" Then
      Call InitCommandLayout(glbGuiConfigs.MasterCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitCommandLayout(glbGuiConfigs.MainCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitCommandLayout(glbGuiConfigs.StockCommandMenuItems)
'   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
'
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
      Call InitCommandLayout(glbGuiConfigs.LedgerCommandMenuItems)
'   ElseIf Node.Key = ROOT_TREE & " 1-6" Then
'   ElseIf Node.Key = ROOT_TREE & " 1-7" Then
'      Call InitCommandLayout(glbGuiConfigs.PackageCommandMenuItems)
'   ElseIf Node.Key = ROOT_TREE & " 1-8" Then
'      Call InitCommandLayout(glbGuiConfigs.HRCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-9" Then
      Call InitCommandLayout(glbGuiConfigs.ProdCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-10" Then
      Call InitCommandLayout(glbGuiConfigs.GLCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-11" Then
      Call InitCommandLayout(glbGuiConfigs.CMCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-12" Then
      Call InitCommandLayout(glbGuiConfigs.PackageCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-13" Then
      Call InitCommandLayout(glbGuiConfigs.TagetCommandMenuItems)
   ElseIf Node.Key = ROOT_TREE & " 1-14" Then
      Call InitCommandLayout(glbGuiConfigs.CostCommandMenuItems)
   End If
   
   Call cmdGeneric(1).SetFocus
End Sub

Private Sub InitCommandLayout(Col As Collection)
Dim D As CMenuItem
Dim Top As Long
Dim Left As Long
Dim I As Long
Dim hight As Long
   Top = cmdGeneric(0).Top
   Left = cmdGeneric(0).Left
   fraGeneric.Height = 1450
   hight = fraGeneric.Height
   For I = 1 To (cmdGeneric.Count - 1)
      cmdGeneric(I).Visible = False
      Unload cmdGeneric(I)
      fraGeneric.Visible = False
   Next I
   
   I = 0
   For Each D In Col
      I = I + 1
      
      Load cmdGeneric(I)
      cmdGeneric(I).Visible = False
      cmdGeneric(I).Picture = LoadPicture(glbParameterObj.MainButton)

      cmdGeneric(I).Left = Left
      cmdGeneric(I).Top = Top
      cmdGeneric(I).Tag = D.KEYWORD
      Call InitMainButton(cmdGeneric(I), D.MENU_TEXT)
      cmdGeneric(I).Visible = True
      fraGeneric.Height = hight
      fraGeneric.Visible = True
      hight = hight + cmdGeneric(0).Height + 10
      Top = Top + cmdGeneric(0).Height + 10
   Next D
     
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 49 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(2))
   ElseIf Shift = 0 And KeyCode = 50 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(3))
   ElseIf Shift = 0 And KeyCode = 51 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(4))
   ElseIf Shift = 0 And KeyCode = 52 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(5))
   ElseIf Shift = 0 And KeyCode = 53 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(6))
   ElseIf Shift = 0 And KeyCode = 54 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(7))
   ElseIf Shift = 0 And KeyCode = 55 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(8))
   ElseIf Shift = 0 And KeyCode = 56 Then
      KeyCode = 0
      Call trvMain_NodeClick(trvMain.Nodes(9))
   End If
End Sub
Private Function SeparateInventorySubTypeColl(TempCollection As Collection, DocType As Long)
Dim MenuSelected As Long
Dim oMenu As CPopupMenu
Dim Mr As CMasterRef
Dim MI As CMenuItem
   
   Set TempCollection = Nothing
   Set TempCollection = New Collection
   For Each Mr In InventorySubTypecoll
      If Mr.INDEX_LINK = DocType Then
         Set MI = New CMenuItem
         MI.KEY_ID = Mr.KEY_ID
         MI.KEYWORD = Mr.KEY_NAME
         Call TempCollection.add(MI)
         Set MI = Nothing
      End If
   Next Mr
         
   If TempCollection.Count > 0 Then
      Set MI = New CMenuItem
      MI.KEY_ID = -1
      MI.KEYWORD = "ทั้งหมด"
      Call TempCollection.add(MI)
   End If
End Function
Private Sub GenerateBillingProgram(Col As Collection)
Dim MI As CMenuItem
Dim TempCol As Collection

   Set Col = Nothing
   Set Col = New Collection
   
   Set MI = New CMenuItem
   MI.KEY_ID = 1
   MI.KEYWORD = "คัดลอกของมูลจากเอกสารใบสั่งซื้อมาเป็นบิลขาย ตามวันที่ส่งของ"
   Call Col.add(MI)
   Set MI = Nothing
End Sub

