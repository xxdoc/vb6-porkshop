VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTagetJobDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTagetJobDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   4895
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOutputType 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1800
      End
      Begin Rebate.uctlTextBox txtOutputAmount 
         Height          =   435
         Left            =   2760
         TabIndex        =   2
         Top             =   1320
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtBatchNo 
         Height          =   435
         Left            =   2760
         TabIndex        =   0
         Top             =   360
         Width           =   360
         _ExtentX        =   3175
         _ExtentY        =   767
      End
      Begin VB.Label lblBatchNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label lblOutputType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   885
         TabIndex        =   3
         Top             =   1950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetJobDetail.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2535
         TabIndex        =   4
         Top             =   1950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetJobDetail.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4185
         TabIndex        =   5
         Top             =   1950
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblOutputAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditTagetJobDetail"
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
Public ParentForm As Form
Private Sub cboOutputType_Click()
   m_HasModify = True
End Sub
Private Sub cboOutputType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
   
   Call InitNormalLabel(lblOutputType, MapText("ประเภทผลิต"))
   Call InitNormalLabel(lblOutputAmount, MapText("ยอดผลิต"))
   Call InitNormalLabel(lblBatchNo, MapText("แบ็ต"))
   
   Call txtOutputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtBatchNO.SetTextLenType(TEXT_INTEGER, glbSetting.FLAG_TYPE)
   
   Call InitCombo(cboOutputType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CTagetJobDetail
         
         Set BD = TempCollection.Item(ID)
         
         cboOutputType.ListIndex = IDToListIndex(cboOutputType, BD.OUTPUT_TYPE_ID)
         txtOutputAmount.Text = BD.OUTPUT_AMOUNT
         txtBatchNO.Text = BD.BATCH_NO
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdNext_Click()
Dim NewID As Long

   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.RefreshGrid
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      cboOutputType.ListIndex = -1
      txtOutputAmount.Text = ""
      txtBatchNO.Text = ""
   End If
   Call QueryData(True)
   
   Call txtBatchNO.SetFocus
   
   Call ParentForm.RefreshGrid
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
Dim I As Long
      
   If Not VerifyTextControl(lblBatchNo, txtBatchNO, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblOutputType, cboOutputType, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblOutputAmount, txtOutputAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim CheckBd As CTagetJobDetail
   For Each CheckBd In TempCollection
      I = I + 1
      
      If CheckBd.BATCH_NO = Val(txtBatchNO.Text) And CheckBd.OUTPUT_TYPE_ID = cboOutputType.ItemData(Minus2Zero(cboOutputType.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล แบ็ต ") & CheckBd.BATCH_NO & "  ประเภท " & cboOutputType.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         cboOutputType.SetFocus
         Exit Function
      End If
      
   Next CheckBd
   
   Dim BD As CTagetJobDetail
   If ShowMode = SHOW_ADD Then
      Set BD = New CTagetJobDetail
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.OUTPUT_TYPE_ID = cboOutputType.ItemData(Minus2Zero(cboOutputType.ListIndex))
   BD.OUTPUT_DESC = cboOutputType.Text
   BD.OUTPUT_AMOUNT = Val(txtOutputAmount.Text)
   BD.BATCH_NO = Val(txtBatchNO.Text)
   
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboOutputType, , , , MASTER_PRODUCTION_TYPE)
      
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
Private Sub txtBatchNO_Change()
   m_HasModify = True
End Sub

Private Sub txtOutputAmount_Change()
   m_HasModify = True
End Sub
