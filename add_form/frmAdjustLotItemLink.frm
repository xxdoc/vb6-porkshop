VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjustLotItemLink 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAdjustLotItemLink.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   6588
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboStockType 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   5175
      End
      Begin VB.ComboBox cboStockGroup 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   5175
      End
      Begin VB.ComboBox cboLocation 
         Height          =   315
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2205
         Width           =   5175
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   2715
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1980
         TabIndex        =   6
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlTextBox txtFromStockNO 
         Height          =   465
         Left            =   1980
         TabIndex        =   0
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlTextBox txtToStockNo 
         Height          =   465
         Left            =   5460
         TabIndex        =   1
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin VB.Label lblStockType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label lblStockGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   1275
         Width           =   1575
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblToStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label lblFromStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1605
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   7
         Top             =   3060
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustLotItemLink.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   13
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   3240
         Width           =   1605
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9495
         TabIndex        =   8
         Top             =   3060
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAdjustLotItemLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private DocLinkLot As Collection
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean

   Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   Status = AdjustLotItemLink
   
   Me.Enabled = True
   
   If Status Then
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
      glbErrorLog.ShowUserError
   Else
      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
      glbErrorLog.ShowUserError
   End If
   
   Call cmdOK_Click
   Exit Sub
   
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Set DocLinkLot = New Collection
      
      Call LoadMaster(cboLocation, Nothing, , , MASTER_LOCATION)
      Call LoadMaster(cboStockGroup, Nothing, , , MASTER_STOCKGROUP)
      Call LoadMaster(cboStockType, Nothing, , , MASTER_STOCKTYPE)
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
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
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblFromStockNo, "จากรหัสสินค้า")
   Call InitNormalLabel(lblToStockNo, "ถึงรหัสสินค้า")
   Call InitNormalLabel(lblLocation, "คลัง")
   Call InitNormalLabel(lblStockGroup, "กลุ่มวัตถุดิบ")
   Call InitNormalLabel(lblStockType, "ประเภทวัตถุดิบ")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFromStockNO.SetKeySearch("STOCK_NO")
   Call txtToStockNo.SetKeySearch("STOCK_NO")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call InitCombo(cboLocation)
   Call InitCombo(cboStockGroup)
   Call InitCombo(cboStockType)
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Function AdjustLotItemLink() As Boolean
Dim Lt As CLotItem
Dim BalanceCollection As Collection
Dim m_LotItem As CLotItem
Dim m_Rs  As ADODB.Recordset
Dim ItemCount As Long
Dim TempRs As ADODB.Recordset
Dim PrevKey1  As String
Dim PrevKey2  As String
Dim UpdateRs As Boolean
Dim OldSubLotAmount As Double
Dim I As Long
   
   Set BalanceCollection = New Collection
   Set m_Rs = New ADODB.Recordset
   Set TempRs = New ADODB.Recordset
   PrevKey1 = ""
   PrevKey2 = ""
   
   ' Load ยอดคงเหลือ
   'Call LoadLeftAmountLotItem(BalanceCollection, -1, -1, cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex)), txtFromStockNO.Text, txtToStockNo.Text)
   Call GetDocItemIDLinkLotItemID(DocLinkLot, cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex)), txtFromStockNO.Text, txtToStockNo.Text)
   
   MasterInd = "22"
   Set m_LotItem = New CLotItem
   m_LotItem.FROM_STOCK_NO = txtFromStockNO.Text
   m_LotItem.TO_STOCK_NO = txtToStockNo.Text
   m_LotItem.LOCATION_ID = cboLocation.ItemData(Minus2Zero(cboLocation.ListIndex))
   m_LotItem.PART_TYPE = cboStockType.ItemData(Minus2Zero(cboStockType.ListIndex))
   m_LotItem.PART_GROUP = cboStockGroup.ItemData(Minus2Zero(cboStockGroup.ListIndex))
   m_LotItem.COUNT_AMOUNT = "Y"
   Call m_LotItem.QueryData(22, m_Rs, ItemCount, True)
   
   I = 0
   prgProgress.Min = 0
   If ItemCount > 0 Then
      prgProgress.Max = ItemCount
   End If
   
   While Not m_Rs.EOF
      Call m_LotItem.PopulateFromRS(22, m_Rs)
      
      I = I + 1
      prgProgress.Value = I
      txtPercent.Text = MyDiffEx(I, ItemCount) * 100
      Me.Refresh
   
      If PrevKey1 <> Trim(Str(m_LotItem.PART_ITEM_ID)) Or PrevKey2 <> Trim(Str(m_LotItem.LOCATION_ID)) Then
         'Set LT = GetObject("CLotItem", BalanceCollection, Trim(m_LotItem.LOCATION_ID & "-" & m_LotItem.PART_ITEM_ID))
         UpdateRs = True
         PrevKey1 = Trim(Str(m_LotItem.PART_ITEM_ID))
         PrevKey2 = Trim(Str(m_LotItem.LOCATION_ID))
         OldSubLotAmount = 0
      Else
         UpdateRs = False
      End If
      If m_LotItem.DOCUMENT_TYPE = EXPORT_DOCTYPE Or m_LotItem.DOCUMENT_TYPE = ADJUST_DOCTYPE Or m_LotItem.DOCUMENT_TYPE = TRANSFER_DOCTYPE Or m_LotItem.DOCUMENT_TYPE = 5 Then
         Call GenerateAutoLotLink(m_LotItem.LOT_ITEM_ID, m_LotItem.TX_AMOUNT, m_LotItem.PART_ITEM_ID, m_LotItem.LOCATION_ID, TempRs, UpdateRs, OldSubLotAmount, False)
      ElseIf m_LotItem.DOCUMENT_TYPE = 10 Or m_LotItem.DOCUMENT_TYPE = 21 Then 'ใบส่งของ กับ ใบขายสด
         Call GenerateAutoLotLink(m_LotItem.LOT_ITEM_ID, m_LotItem.TX_AMOUNT, m_LotItem.PART_ITEM_ID, m_LotItem.LOCATION_ID, TempRs, UpdateRs, OldSubLotAmount, True)
      Else
         ''debug.print
      End If
      m_Rs.MoveNext
   Wend
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set BalanceCollection = Nothing
   Set m_LotItem = Nothing
   Set DocLinkLot = Nothing
   AdjustLotItemLink = True
   MasterInd = "1"
End Function
Private Sub Form_Unload(Cancel As Integer)
   Set DocLinkLot = Nothing
End Sub
