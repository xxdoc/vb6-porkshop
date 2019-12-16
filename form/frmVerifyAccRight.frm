VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmVerifyAccRight 
   Caption         =   "Form1"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   1125
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2566
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextBox txtUsername 
         Height          =   435
         Left            =   2040
         TabIndex        =   0
         Top             =   120
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtPassword 
         Height          =   435
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   767
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmVerifyAccRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AccName As String
Public GrantRight As Boolean
Private m_ADOConn As ADODB.Connection
Private Sub Form_Activate()
   ''debug.print
   GrantRight = False
End Sub

Private Sub Form_Load()
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("กรุณากรอกชื่อผู้ใช้และรหัสผ่าน")
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Call InitNormalLabel(lblUsername, MapText("ชื่อผู้ใช้"))
   Call InitNormalLabel(lblPassword, MapText("รหัสผ่าน"))
   
   Call txtUsername.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtUsername.SetTextType(1)
   Call txtPassword.SetTextLenType(TEXT_STRING, glbSetting.PASSWORD_TYPE)
   txtPassword.PasswordChar = "*"
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub
Private Sub txtPassword_LostFocus()
   If Not VerifyTextControl(lblUsername, txtUsername, False) Then
      Exit Sub
   End If
   
   If Not VerifyTextControl(lblPassword, txtPassword, False) Then
      Exit Sub
   End If
   
   Call CheckAccRightUserPassword
   
   Unload Me
End Sub
Private Function CheckAccRightUserPassword() As Boolean
Dim m_Rs1 As ADODB.Recordset
Dim ItemCount  As Long
Dim SQL1 As String
Dim ErrorObj As clsErrorLog
   Set m_Rs1 = New ADODB.Recordset
   Set ErrorObj = New clsErrorLog
   
   SQL1 = "SELECT UA.*, UG.*,GR.*,RI.* "
   SQL1 = SQL1 & " FROM USER_ACCOUNT UA,USER_GROUP UG,GROUP_RIGHT GR,RIGHT_ITEM RI "
   SQL1 = SQL1 & "WHERE (UA.GROUP_ID = UG.GROUP_ID) "
   SQL1 = SQL1 & "AND (UG.GROUP_ID = GR.GROUP_ID) "
   SQL1 = SQL1 & "AND (GR.RIGHT_ID = RI.RIGHT_ID) "
   SQL1 = SQL1 & "AND (RI.RIGHT_ITEM_NAME = '" & ChangeQuote(AccName) & "' ) "
   SQL1 = SQL1 & "AND (UA.USER_NAME = '" & ChangeQuote(txtUsername.Text) & "' ) "
   SQL1 = SQL1 & "AND (UA.USER_PASSWORD = '" & ChangeQuote(EncryptText(txtPassword.Text)) & "' ) "
   SQL1 = SQL1 & "AND (GR.RIGHT_STATUS = 'Y' ) "
   
   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs1, ErrorObj) Then
      Exit Function
   End If
   
   If (m_Rs1.EOF) Or (NVLS(m_Rs1("USER_STATUS"), "Y") <> "Y") Then
      ErrorObj.LocalErrorMsg = "บัญชีรายชื่อนี้ไม่สามารถเข้าถึงข้อมูลส่วนนี้ได้"
      ErrorObj.SystemErrorMsg = " ไม่สามารถเข้าถึงส่วน " & AccName
      ErrorObj.RoutineName = "CheckAccRightUserPassword"
      ErrorObj.ModuleName = "frmVerifyAccRight"
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      GrantRight = False
      Exit Function
   End If
   
   Set m_Rs1 = Nothing
   Set ErrorObj = Nothing
   GrantRight = True
End Function
