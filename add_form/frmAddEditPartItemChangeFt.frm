VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPartItemChangeFt 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPartItemChangeFt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4895
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Rebate.uctlTextBox txtBox 
         Height          =   435
         Left            =   2280
         TabIndex        =   0
         Top             =   240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtTray 
         Height          =   435
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtPack 
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   1200
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
      End
      Begin VB.Label lblPack 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label lblTray 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Width           =   1725
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   855
         TabIndex        =   3
         Top             =   1950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItemChangeFt.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   2505
         TabIndex        =   4
         Top             =   1950
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblBox 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditPartItemChangeFt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form
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
   
   Call InitNormalLabel(lblBox, MapText("กล่อง"))
   Call InitNormalLabel(lblTray, MapText("ถาด"))
   Call InitNormalLabel(lblPack, MapText("แพ็ค"))
   
   Call txtBox.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTray.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPack.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
      
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   'cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   'Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CStockCodeChangeFt
         
         Set BD = TempCollection.Item(ID)
         
         txtBox.Text = BD.BOX
         txtTray.Text = BD.TRAY
         txtPack.Text = BD.PACK
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub
'Private Sub cmdNext_Click()
'Dim NewID As Long
'
'   If Not SaveData Then
'      Exit Sub
'   End If
'
'   If ShowMode = SHOW_EDIT Then
'      NewID = GetNextID(ID, TempCollection)
'      If ID = NewID Then
'         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
'         glbErrorLog.ShowUserError
'
'         Call ParentForm.RefreshGrid("UNIT_CHANGE_VERIFY")
'         Exit Sub
'      End If
'
'      ID = NewID
'   ElseIf ShowMode = SHOW_ADD Then
'      txtBox.Text = ""
'      txtTray.Text = ""
'      txtPack.Text = ""
'   End If
'   Call QueryData(True)
'
'   Call txtBox.SetFocus
'
'   Call ParentForm.RefreshGrid("UNIT_CHANGE_VERIFY")
'End Sub
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
Dim I As Long
          
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim BD As CStockCodeChangeFt
   If ShowMode = SHOW_ADD Then
      Set BD = New CStockCodeChangeFt
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.BOX = Val(txtBox.Text)
   BD.TRAY = Val(txtTray.Text)
   BD.PACK = Val(txtPack.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
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
End Sub

Private Sub txtBox_Change()
   m_HasModify = True
End Sub

Private Sub txtPack_Change()
   m_HasModify = True
End Sub

Private Sub txtTray_Change()
   m_HasModify = True
End Sub
