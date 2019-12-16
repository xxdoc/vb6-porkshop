VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditGLAcc 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmAddEditGLAcc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSPanel Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   -90
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1296
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame Frame1 
      Height          =   2685
      Left            =   0
      TabIndex        =   5
      Top             =   540
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4736
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTextBox txtName 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1230
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboParent 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   3855
      End
      Begin Xivess.uctlTextBox txtLedgerCode 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         Top             =   780
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4065
         TabIndex        =   4
         Top             =   1860
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2415
         TabIndex        =   3
         Top             =   1860
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditGLAcc.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblLedgerCode 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         TabIndex        =   9
         Top             =   780
         Width           =   1665
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         TabIndex        =   8
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label lblParent 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   330
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmAddEditGLAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmAddEditCustomerInfo"

Private HasActivate As Boolean
Private HasModify As Boolean
Public HeaderText As String
Public OKClick As Boolean
Public OrganizeID As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_Rs As ADODB.Recordset

Private m_GLAcc As CGLAccount
Private m_GLAccs As Collection

Private Sub cboGroup_Click()
   HasModify = True
End Sub

Private Sub chkStatus_Click()
   HasModify = True
End Sub

Private Sub cboContactType_Click()
   HasModify = True
End Sub

Private Sub cboQualifier_Click()
   HasModify = True
End Sub

Private Sub cboDependency_Click()
   HasModify = True
End Sub

Private Sub cboCardType_Click()
   HasModify = True
End Sub

Private Sub cboParent_Click()
   HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorHandler
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("GL_ACC_ADD", "") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   Else
      If Not VerifyAccessRight("GL_ACC_EDIT", "") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Activate"
   
   If Not VerifyCombo(lblParent, cboParent, True) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblName, txtName) Then
      Exit Sub
   End If
   
   If Not HasModify Then
      Unload Me
      Exit Sub
   End If

   Call m_GLAcc.SetFieldValue("ACC_NAME", txtName.Text)
   Call m_GLAcc.SetFieldValue("ACC_CODE", txtLedgerCode.Text)
   If cboParent.ListIndex > 0 Then
      Call m_GLAcc.SetFieldValue("PARENT_ID", cboParent.ItemData(cboParent.ListIndex))
   Else
      Call m_GLAcc.SetFieldValue("PARENT_ID", -1)
   End If

   Call EnableForm(Me, False)
   m_GLAcc.ShowMode = ShowMode
   If Not glbDaily.AddEditGLAccount(m_GLAcc, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim IsOK As Boolean

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"

   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
      
      Call m_GLAcc.SetFieldValue("GL_ACCOUNT_ID", -1)
      Call LoadGLAccount(m_GLAcc, cboParent, m_GLAccs)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call EnableForm(Me, False)
         Call m_GLAcc.SetFieldValue("GL_ACCOUNT_ID", OrganizeID)
         If Not glbDaily.QueryGLAccount(m_GLAcc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If ItemCount > 0 Then
            Call m_GLAcc.PopulateFromRS(1, m_Rs)
            
            txtName.Text = m_GLAcc.GetFieldValue("ACC_NAME")
            txtLedgerCode.Text = m_GLAcc.GetFieldValue("ACC_CODE")
            cboParent.ListIndex = IDToListIndex(cboParent, m_GLAcc.GetFieldValue("PARENT_ID"))
         End If
         Call EnableForm(Me, True)
         HasModify = False
      End If
   End If
   
   Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      MsgBox Me.Name
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

   Frame2.BackColor = GLB_FORM_COLOR
   Frame2.Font.Name = GLB_FONT
   Frame2.Font.Bold = True
   Frame2.Font.Size = 19
   
   Frame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   Frame2.BackColor = GLB_HEAD_COLOR
   Frame2.Caption = HeaderText
   
   OKClick = False
   HasActivate = False

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))

   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitNormalLabel(lblName, MapText("ชื่อบัญชี"))
   Call InitNormalLabel(lblLedgerCode, MapText("รหัสบัญชี"))
   Call InitNormalLabel(lblParent, MapText("บัญชีต้นสังกัด"))
   
   Call txtLedgerCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboParent)
   
   Call EnableForm(Me, True)
   HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_GLAcc = New CGLAccount
   Set m_GLAccs = New Collection
   
   HasActivate = False
   
   Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_GLAcc = Nothing
   Set m_GLAccs = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Sub txtDesc_Change()
   HasModify = True
End Sub

Private Sub txtLedgerCode_Change()
   HasModify = True
End Sub

Private Sub txtName_Change()
   HasModify = True
End Sub
