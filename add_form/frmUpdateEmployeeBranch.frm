VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateEmployeeBranch 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15465
   Icon            =   "frmUpdateEmployeeBranch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   15465
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3555
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   6271
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   2220
         TabIndex        =   2
         Top             =   1920
         Width           =   12915
         _ExtentX        =   22781
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   15675
         _ExtentX        =   27649
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtPercent 
         Height          =   465
         Left            =   2220
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin PorkShop.uctlTextLookup uctlEmployeeLookUp 
         Height          =   435
         Left            =   2220
         TabIndex        =   10
         Top             =   960
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlToEmployeeLookUp 
         Height          =   435
         Left            =   9900
         TabIndex        =   12
         Top             =   960
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblToEmployee 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   8160
         TabIndex        =   9
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   480
         TabIndex        =   11
         Top             =   1020
         Width           =   1575
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   8160
         TabIndex        =   0
         Top             =   2580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmUpdateEmployeeBranch.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   4080
         TabIndex        =   8
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   480
         TabIndex        =   6
         Top             =   2400
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9855
         TabIndex        =   1
         Top             =   2580
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmUpdateEmployeeBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public HeaderText As String

Private Emp As CEmployee
Private EmpTo As CEmployee

Private Sub cmdStart_Click()
Dim Status As Boolean
Dim IsOK As Boolean

   If Not VerifyCombo(lblEmployee, uctlEmployeeLookUp.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblToEmployee, uctlToEmployeeLookUp.MyCombo, False) Then
      Exit Sub
   End If
   
   If uctlEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookUp.MyCombo.ListIndex)) = uctlToEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlToEmployeeLookUp.MyCombo.ListIndex)) Then
      glbErrorLog.LocalErrorMsg = MapText("พนักงานขายซ้ำ")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   Status = AdjustEmployee
   
   Me.Enabled = True
   
   If Status Then
      If ConfirmSave Then
         Call glbDaily.CommitTransaction
         glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
         glbErrorLog.ShowUserError
      Else
         Call glbDaily.RollbackTransaction
         glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
         glbErrorLog.ShowUserError
      End If
   Else
      Call glbDaily.RollbackTransaction
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
      
      Call LoadEmployee(Emp, uctlEmployeeLookUp.MyCombo)
      Set uctlEmployeeLookUp.MyCollection = m_EmployeeColl
      
      Call LoadEmployee(EmpTo, uctlToEmployeeLookUp.MyCombo)
      Set uctlToEmployeeLookUp.MyCollection = m_EmployeeColl
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
   
   Call InitNormalLabel(lblEmployee, "จากพนักงานขาย", RGB(255, 0, 0))
   Call InitNormalLabel(lblToEmployee, "ไปยังพนักงานขาย", RGB(255, 0, 0))
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
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
   Call EnableForm(Me, False)
   
   Set Emp = New CEmployee
   Set EmpTo = New CEmployee
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Function AdjustEmployee() As Boolean
Dim Mr As CMasterRef
Dim IDFrom As Long
Dim IDTo As Long
   
   IDFrom = uctlEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookUp.MyCombo.ListIndex))
   IDTo = uctlToEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlToEmployeeLookUp.MyCombo.ListIndex))
   Set Mr = New CMasterRef
   
   Call Mr.UpdateEmployeeBranch(IDFrom, IDTo)
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   AdjustEmployee = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set Emp = Nothing
   Set EmpTo = Nothing
End Sub

