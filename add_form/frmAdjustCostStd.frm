VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjustCostStd 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAdjustCostStd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7223
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2460
         TabIndex        =   12
         Top             =   3075
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtPercent 
         Height          =   465
         Left            =   2460
         TabIndex        =   13
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlTextBox txtFromStockNO 
         Height          =   465
         Left            =   2460
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlTextBox txtToStockNo 
         Height          =   465
         Left            =   5580
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlDate uctlFromDate 
         Height          =   405
         Left            =   2460
         TabIndex        =   7
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Rebate.uctlDate uctlToDate 
         Height          =   405
         Left            =   7620
         TabIndex        =   8
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Rebate.uctlTextBox txtStdCost 
         Height          =   465
         Left            =   2460
         TabIndex        =   9
         Top             =   2115
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Rebate.uctlTextBox txtStdCostPlus 
         Height          =   465
         Left            =   7620
         TabIndex        =   10
         Top             =   2115
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Threed.SSCheck UpDateToStockCode 
         Height          =   435
         Left            =   2520
         TabIndex        =   11
         Top             =   2620
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblStdCostPlus 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4200
         TabIndex        =   26
         Top             =   2205
         Width           =   3315
      End
      Begin Threed.SSOption SSOption2 
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   131073
         Caption         =   "SSOption1"
      End
      Begin Threed.SSCheck chk3 
         Height          =   435
         Left            =   10440
         TabIndex        =   6
         Top             =   1200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblStdCost 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   2235
      End
      Begin Threed.SSCheck chk2 
         Height          =   435
         Left            =   9120
         TabIndex        =   5
         Top             =   1200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCheck chk1 
         Height          =   435
         Left            =   7680
         TabIndex        =   4
         Top             =   1200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         TabIndex        =   24
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblToStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblFromStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   1320
         Width           =   1605
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   8280
         TabIndex        =   14
         Top             =   3420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustCostStd.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4200
         TabIndex        =   20
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Top             =   3120
         Width           =   1605
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   720
         TabIndex        =   18
         Top             =   3600
         Width           =   1605
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   9975
         TabIndex        =   15
         Top             =   3420
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAdjustCostStd"
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

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub chk1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chk2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chk3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
   
   If Not VerifyTextControl(lblStdCost, txtStdCost, Not (txtStdCost.Enabled)) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblStdCostPlus, txtStdCostPlus, Not (txtStdCostPlus.Enabled)) Then
      Exit Sub
   End If
   
   If chk1.Value = ssCBUnchecked And chk2.Value = ssCBUnchecked And chk3.Value = ssCBUnchecked Then
      glbErrorLog.LocalErrorMsg = "กรุณาใส่ประเภทเอกสารอย่างน้อยหนึ่งประเภท"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   Status = AdjustStdCost
   
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
   Call InitNormalLabel(lblFromDate, "จากวันที่")
   Call InitNormalLabel(lblToDate, "ถึงวันที่")
   Call InitNormalLabel(lblStdCost, "ต้นทุนมาตรฐานใหม่")
   Call InitNormalLabel(lblStdCostPlus, "ส่วนเพิ่ม/ลดต้นทุนมาตรฐาน")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   Call txtStdCost.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtStdCostPlus.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFromStockNO.SetKeySearch("STOCK_NO")
   Call txtToStockNo.SetKeySearch("STOCK_NO")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call InitCheckBox(chk1, "ขายเชื่อ")
   chk1.Value = ssCBChecked
   Call InitCheckBox(chk2, "ขายสด")
   chk2.Value = ssCBChecked
   Call InitCheckBox(chk3, "รับคืน")
   chk3.Value = ssCBChecked
   Call InitCheckBox(UpDateToStockCode, "อัพเดดต้นทุนไปที่วัตถุดิบ")
   UpDateToStockCode.Value = ssCBUnchecked
   
   Call InitOptionEx(SSOption1, "กำหนด ค่ามาตรฐานใหม่")
   Call InitOptionEx(SSOption2, "เพิ่ม/ลด ค่ามาตรฐานใหม่")
   SSOption1.Value = True
   txtStdCost.Enabled = True
   txtStdCostPlus.Enabled = False
   
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
Private Function AdjustStdCost() As Boolean
Dim SQL As String
Dim Lt As CLotItem
Dim BD As CBillingDoc

Dim DocumentTypeSet As String
   
   AdjustStdCost = False
   
   Set BD = New CBillingDoc
   BD.FROM_DATE = uctlFromDate.ShowDate
   BD.TO_DATE = uctlToDate.ShowDate
   BD.FROM_STOCK_NO = txtFromStockNO.Text
   BD.TO_STOCK_NO = txtToStockNo.Text
   BD.CAPITAL_AMOUNT = Val(txtStdCost.Text)
   DocumentTypeSet = GenerateDocumentTypeSet(2)
   Call BD.UpDateStdCost(DocumentTypeSet, Val(txtStdCostPlus.Text), SSOption2.Value)
   Set BD = Nothing
   
   Set Lt = New CLotItem
   Lt.FROM_DOC_DATE = uctlFromDate.ShowDate
   Lt.TO_DOC_DATE = uctlToDate.ShowDate
   Lt.FROM_STOCK_NO = txtFromStockNO.Text
   Lt.TO_STOCK_NO = txtToStockNo.Text
   Lt.AVG_PRICE = Val(txtStdCost.Text)
   DocumentTypeSet = GenerateDocumentTypeSet(1)
   If UpDateToStockCode.Value = ssCBChecked Then
      Call Lt.UpDateStdCost(DocumentTypeSet, Val(txtStdCostPlus.Text), True, SSOption2.Value)
   Else
      Call Lt.UpDateStdCost(DocumentTypeSet, Val(txtStdCostPlus.Text), False, SSOption2.Value)
   End If
   
   Set Lt = Nothing
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   
   AdjustStdCost = True
   
End Function
Public Function GenerateDocumentTypeSet(GType As Byte) As String
   If GType = 1 Then
      GenerateDocumentTypeSet = "("
      If chk1.Value = ssCBChecked Then
         GenerateDocumentTypeSet = GenerateDocumentTypeSet & "10,"
      End If
      If chk2.Value = ssCBChecked Then
         GenerateDocumentTypeSet = GenerateDocumentTypeSet & "21,"
      End If
      If chk3.Value = ssCBChecked Then
         GenerateDocumentTypeSet = GenerateDocumentTypeSet & "30,"
      End If
      GenerateDocumentTypeSet = Left(GenerateDocumentTypeSet, Len(GenerateDocumentTypeSet) - 1) & ")"
   ElseIf GType = 2 Then
      GenerateDocumentTypeSet = "("
      If chk1.Value = ssCBChecked Then
         GenerateDocumentTypeSet = GenerateDocumentTypeSet & "3,"
      End If
      If chk2.Value = ssCBChecked Then
         GenerateDocumentTypeSet = GenerateDocumentTypeSet & "4,"
      End If
      If chk3.Value = ssCBChecked Then
         GenerateDocumentTypeSet = GenerateDocumentTypeSet & "6,"
      End If
      GenerateDocumentTypeSet = Left(GenerateDocumentTypeSet, Len(GenerateDocumentTypeSet) - 1) & ")"
   End If
End Function
Private Sub SSOption1_Click(Value As Integer)
   If SSOption1.Value Then
      txtStdCost.Enabled = True
      txtStdCostPlus.Enabled = False
   Else
      txtStdCost.Enabled = False
      txtStdCostPlus.Enabled = True
   End If
End Sub

Private Sub SSOption1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
   If SSOption2.Value Then
      txtStdCost.Enabled = False
      txtStdCostPlus.Enabled = True
   Else
      txtStdCost.Enabled = True
      txtStdCostPlus.Enabled = False
   End If
End Sub
Private Sub SSOption2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub UpDateToStockCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
