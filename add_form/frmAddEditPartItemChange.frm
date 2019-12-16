VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPartItemChange 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPartItemChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   5318
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnit 
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
         TabIndex        =   0
         Top             =   240
         Width           =   1800
      End
      Begin Rebate.uctlTextBox txtUnitAmount 
         Height          =   435
         Left            =   2760
         TabIndex        =   1
         Top             =   720
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlAparMasLookup 
         Height          =   435
         Left            =   2760
         TabIndex        =   9
         Top             =   1320
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblAparMas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   1245
         TabIndex        =   2
         Top             =   2190
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItemChange.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2895
         TabIndex        =   3
         Top             =   2190
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItemChange.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4545
         TabIndex        =   4
         Top             =   2190
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnitAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditPartItemChange"
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

Public ParentUnitID As Long
Public ChildUnitID As Long
Private Sub cboUnit_Click()
   m_HasModify = True
End Sub
Private Sub cboUnit_KeyPress(KeyAscii As Integer)
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
   
   Call InitNormalLabel(lblUnit, MapText("หน่วย"))
   Call InitNormalLabel(lblUnitAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblAparMas, MapText("รหัสลูกค้า"))
   
   Call txtUnitAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboUnit)
   
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
         Dim BD As CStockCodeChange
         
         Set BD = TempCollection.Item(ID)
         
         cboUnit.ListIndex = IDToListIndex(cboUnit, BD.UNIT_CHANGE_ID)
         txtUnitAmount.Text = BD.UNIT_CHANGE_AMOUNT
         uctlAparMasLookup.MyCombo.ListIndex = IDToListIndex(uctlAparMasLookup.MyCombo, BD.CUSTOMER_ID)
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
         
         Call ParentForm.RefreshGrid("UNIT_CHANGE")
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
      cboUnit.ListIndex = -1
      txtUnitAmount.Text = ""
   End If
   Call QueryData(True)
   
   Call cboUnit.SetFocus
   
   Call ParentForm.RefreshGrid("UNIT_CHANGE")
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
      
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
      
   If Not VerifyTextControl(lblUnitAmount, txtUnitAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If (ParentUnitID = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))) Or (ChildUnitID = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาใส่ หน่วย ให้ไม่ตรงกับข้อมูล หน่วยใหญ่ และ หน่วยย่อย")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Dim CheckBd As CStockCodeChange
   For Each CheckBd In TempCollection
      I = I + 1

      If CheckBd.UNIT_CHANGE_ID = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex)) And CheckBd.CUSTOMER_ID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & cboUnit.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   
   Next CheckBd
   
   Dim BD As CStockCodeChange
   If ShowMode = SHOW_ADD Then
      Set BD = New CStockCodeChange
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.UNIT_CHANGE_ID = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   BD.UNIT_CHANGE_NAME = cboUnit.Text
   BD.UNIT_CHANGE_AMOUNT = Val(txtUnitAmount.Text)
   BD.CUSTOMER_ID = uctlAparMasLookup.MyCombo.ItemData(Minus2Zero(uctlAparMasLookup.MyCombo.ListIndex))
   BD.CUSTOMER_CODE = uctlAparMasLookup.MyTextBox.Text
   BD.CUSTOMER_NAME = uctlAparMasLookup.MyCombo.Text
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboUnit, , , , MASTER_UNIT)
      
      Dim m_Apm  As CAPARMas
      Set m_Apm = New CAPARMas
      Call LoadApArMas(m_Apm, uctlAparMasLookup.MyCombo)
      Set uctlAparMasLookup.MyCollection = m_CustomerColl
      Set m_Apm = Nothing
      
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
Private Sub txtUnitAmount_Change()
   m_HasModify = True
End Sub
Private Sub uctlApArMasLookup_Change()
   m_HasModify = True
End Sub
