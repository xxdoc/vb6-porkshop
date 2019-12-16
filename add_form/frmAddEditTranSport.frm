VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTranSport 
   BackColor       =   &H80000000&
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   Icon            =   "frmAddEditTranSport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8565
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   15108
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   2400
         Width           =   11055
      End
      Begin VB.ComboBox cboDriver 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   2235
      End
      Begin VB.ComboBox cboCarLicense 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1515
      End
      Begin VB.ComboBox cboTransportor 
         Height          =   315
         Left            =   8340
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   960
         Width           =   2835
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtTranSportPath 
         Height          =   435
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtCarType 
         Height          =   435
         Left            =   1800
         TabIndex        =   4
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtCostPerRound 
         Height          =   435
         Left            =   8940
         TabIndex        =   5
         Top             =   1920
         Width           =   2230
         _ExtentX        =   3942
         _ExtentY        =   767
      End
      Begin VB.Label lblCostPerRound 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   7440
         TabIndex        =   14
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label lblCarType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label lblTranSportPath 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   1530
         Width           =   1455
      End
      Begin VB.Label lblCarLicense 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4410
         TabIndex        =   13
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblTransportor 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   6960
         TabIndex        =   12
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblDriver 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1020
         Width           =   1455
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   1755
         TabIndex        =   8
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   105
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTranSport.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTranSport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_TranSportDetail As CTranSportDetail

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private Sub cboCarLicense_Click()
   m_HasModify = True
End Sub

Private Sub cboDriver_Click()
   m_HasModify = True
End Sub
Private Sub cboTransportor_Click()
   m_HasModify = True
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_TranSportDetail.TRANSPORT_DETAIL_ID = ID
      
      Call m_TranSportDetail.QueryData(1, m_Rs, ItemCount, True)
   End If
   
   If ItemCount > 0 Then
      Call m_TranSportDetail.PopulateFromRS(1, m_Rs)
      
      cboDriver.ListIndex = IDToListIndex(cboDriver, m_TranSportDetail.DRIVER_ID)
      cboCarLicense.ListIndex = IDToListIndex(cboCarLicense, m_TranSportDetail.CAR_LICENSE_ID)
      cboTranSportor.ListIndex = IDToListIndex(cboTranSportor, m_TranSportDetail.TRANSPORTOR_ID)
      txtDesc.Text = m_TranSportDetail.TRANSPORT_DETAIL_DESC
      
      txtTranSportPath.Text = m_TranSportDetail.TRANSPORT_PATH
      txtCarType.Text = m_TranSportDetail.CAR_TYPE
      txtCostPerRound.Text = m_TranSportDetail.COST_PER_ROUND
      
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_TranSportDetail.ShowMode = ShowMode
   m_TranSportDetail.TRANSPORT_DETAIL_ID = ID
   m_TranSportDetail.DRIVER_ID = cboDriver.ItemData(Minus2Zero(cboDriver.ListIndex))
   m_TranSportDetail.CAR_LICENSE_ID = cboCarLicense.ItemData(Minus2Zero(cboCarLicense.ListIndex))
   m_TranSportDetail.TRANSPORTOR_ID = cboTranSportor.ItemData(Minus2Zero(cboTranSportor.ListIndex))
   m_TranSportDetail.TRANSPORT_DETAIL_DESC = txtDesc.Text
   
   m_TranSportDetail.TRANSPORT_PATH = txtTranSportPath.Text
   m_TranSportDetail.CAR_TYPE = txtCarType.Text
   m_TranSportDetail.COST_PER_ROUND = Val(txtCostPerRound.Text)
   
   If m_TranSportDetail.DRIVER_ID <= 0 And m_TranSportDetail.CAR_LICENSE_ID <= 0 And m_TranSportDetail.TRANSPORTOR_ID <= 0 Then
      glbErrorLog.LocalErrorMsg = "กรุณาใส่อย่างน้อย 1 รายการ"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not CheckUniqueNs(TRANSPORT_DETAIL, m_TranSportDetail.DRIVER_ID, ID, Trim(Str(m_TranSportDetail.CAR_LICENSE_ID)), Trim(Str(m_TranSportDetail.TRANSPORTOR_ID)), True) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลคนขับ") & " " & EmptyToString(cboDriver.Text, "N/A") & " และทะเบียน " & EmptyToString(cboCarLicense.Text, "N/A") & " และขนส่ง " & EmptyToString(cboTranSportor.Text, "N/A") & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   
   Call m_TranSportDetail.AddEditData
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadMaster(cboDriver, , , , MASTER_DRIVER)
      Call LoadMaster(cboCarLicense, , , , MASTER_CAR_LICENSE)
      Call LoadMaster(cboTranSportor, , , , MASTER_TRANSPORTOR)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDriver, MapText("คนขับ"))
   Call InitNormalLabel(lblCarLicense, MapText("ทะเบียน"))
   Call InitNormalLabel(lblTranSportor, MapText("ขนส่ง"))
   Call InitNormalLabel(lblTranSportPath, MapText("เส้นทางขนส่ง"))
   Call InitNormalLabel(lblCarType, MapText("ประเภทรถ"))
   Call InitNormalLabel(lblCostPerRound, MapText("ค่าขนส่ง/เที่ยว"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboDriver)
   Call InitCombo(cboCarLicense)
   Call InitCombo(cboTranSportor)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
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
   
   Set m_TranSportDetail = New CTranSportDetail
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Resize()
   pnlHeader.Width = ScaleWidth
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   
   txtDesc.Width = ScaleWidth - txtDesc.Left - 200
   txtDesc.Height = ScaleHeight - 3000
   
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_TranSportDetail = Nothing
End Sub

Private Sub txtCarType_Change()
   m_HasModify = True
End Sub

Private Sub txtCostPerRound_Change()
   m_HasModify = True
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtTranSportPath_Change()
   m_HasModify = True
End Sub
