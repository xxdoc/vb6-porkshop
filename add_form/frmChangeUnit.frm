VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmChangeUnit 
   Caption         =   "Form1"
   ClientHeight    =   690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   Picture         =   "frmChangeUnit.frx":0000
   ScaleHeight     =   690
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1296
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextBox txtMultiple 
         Height          =   415
         Left            =   5760
         TabIndex        =   1
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
      End
      Begin PorkShop.uctlTextLookup uctlUnitChange 
         Height          =   415
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   741
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmChangeUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HeaderText As String

Public UnitID As Long
Public UnitName As String
Public UnitMName As String
Public Multiple As Double
Public m_HasModify As Boolean

Private UnitColls As Collection
Private Mr As CMasterRef

Private Sub Form_Activate()
   
   uctlUnitChange.MyCombo.ListIndex = IDToListIndex(uctlUnitChange.MyCombo, UnitID)
   Call InitNormalLabel(lblUnit, UnitMName)
   txtMultiple.Text = Multiple
   
   m_HasModify = False
End Sub

Private Sub Form_Load()
   
   Set Mr = New CMasterRef
   Set UnitColls = New Collection
   
   Call InitFormLayout
   
End Sub
Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.KeyPreview = True
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   Call txtMultiple.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call LoadMaster(uctlUnitChange.MyCombo, UnitColls, , , MASTER_UNIT)
   Set uctlUnitChange.MyCollection = UnitColls
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Mr = Nothing
   Set UnitColls = Nothing
End Sub

Private Sub txtMultiple_Change()
   m_HasModify = True
End Sub

Private Sub txtMultiple_LostFocus()
Dim ID As Long
   ID = uctlUnitChange.MyCombo.ItemData(Minus2Zero(uctlUnitChange.MyCombo.ListIndex))
   If ID <= 0 Or Val(txtMultiple.Text) <= 0 Then
      Call uctlUnitChange.MyTextBox.SetFocus
   Else
      UnitID = ID
      UnitName = uctlUnitChange.MyCombo.Text
      Multiple = Val(txtMultiple.Text)
      Unload Me
   End If
End Sub

Private Sub uctlUnitChange_Change()
   m_HasModify = True
End Sub
