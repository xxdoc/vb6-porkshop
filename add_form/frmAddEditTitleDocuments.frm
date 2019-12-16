VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTitleDocuments 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   Icon            =   "frmAddEditTitleDocuments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   9375
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2625
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   4630
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtTitle 
         Height          =   435
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   4395
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtSubTitle 
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   767
      End
      Begin VB.Label lblSubTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5640
         TabIndex        =   3
         Top             =   1920
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3480
         TabIndex        =   2
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTitleDocuments.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTitleDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_MasterRef As CMasterRef

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object
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

      Call EnableForm(Me, False)
      Set m_MasterRef = New CMasterRef
      m_MasterRef.KEY_ID = ID
      m_MasterRef.MASTER_AREA = 999
      Call m_MasterRef.QueryData(1, m_Rs, ItemCount, True)
      
      Call m_MasterRef.PopulateFromRS(1, m_Rs)
      If m_MasterRef.KEY_CODE = "A-01" Then
        txtTitle.Text = "ใบลดหนี้รับคืนสินค้า"
      Else
         txtTitle.Text = ""
      End If
      txtSubTitle.Text = m_MasterRef.KEY_NAME
      Call EnableForm(Me, True)
      txtSubTitle.SetFocus
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_MasterRef.ShowMode = SHOW_EDIT
   m_MasterRef.KEY_NAME = txtSubTitle.Text
   
   Call m_MasterRef.AddEditData

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call QueryData(True)
      
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
   ElseIf Shift = 0 And KeyCode = 123 Then
'      Call AddMemoNote
      KeyCode = 0
   End If
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblTitle, MapText("ชื่อเอกสาร"))
   Call InitNormalLabel(lblSubTitle, MapText("เพิ่มเติม"))
   
   Call txtTitle.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   Call txtSubTitle.SetTextLenType(TEXT_STRING, glbSetting.NAME_TYPE)
   txtTitle.Enabled = False
    
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
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
   
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub txtSubTitle_Change()
   m_HasModify = True
End Sub
