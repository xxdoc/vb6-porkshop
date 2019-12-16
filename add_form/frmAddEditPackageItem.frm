VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPackageItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPackageItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   6165
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextBox txtPartCost 
         Height          =   435
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlFromProDate 
         Height          =   405
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtProCost 
         Height          =   435
         Left            =   2160
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
         _ExtentX        =   13361
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlToProDate 
         Height          =   405
         Left            =   6720
         TabIndex        =   3
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblToProDate 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   6120
         TabIndex        =   14
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblProCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblfromProDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblPartLookup 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   3240
         TabIndex        =   5
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackageItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4920
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPackageItem.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6600
         TabIndex        =   7
         Top             =   2640
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartCost 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditPackageItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Public m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_PackageItem As CPackageDetail

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_Products As Collection

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
   
   Call InitNormalLabel(lblPartLookup, MapText("สินค้า"))
   Call InitNormalLabel(lblPartCost, MapText("ราคาขาย"))
   Call InitNormalLabel(lblfromProDate, MapText("ช่วงโปรโมชั่น"))
   Call InitNormalLabel(lblToProDate, MapText("ถึง"))
   Call InitNormalLabel(lblProCost, MapText("ราคาโปร"))
      
   Call txtPartCost.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
    Call txtProCost.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
    
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
         Dim BD As CPackageDetail
         
         Set BD = TempCollection.Item(ID)
         
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, BD.GetFieldValue("PART_ITEM_ID"))
         txtPartCost.Text = BD.GetFieldValue("PART_ITEM_COST")
         
         uctlFromProDate.ShowDate = BD.GetFieldValue("PRO_FROM_DATE")
         uctlToProDate.ShowDate = BD.GetFieldValue("PRO_TO_DATE")
         txtProCost.Text = BD.GetFieldValue("PRO_ITEM_COST")
         
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
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
       uctlPartLookup.MyCombo.ListIndex = -1
      txtPartCost.Text = ""
   End If
   
   Call uctlPartLookup.SetFocus
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

      If Not VerifyCombo(lblPartLookup, uctlPartLookup.MyCombo, False) Then
         Exit Function
      End If
      
   If Not VerifyTextControl(lblPartCost, txtPartCost, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim CheckBd As CPackageDetail
   For Each CheckBd In TempCollection
      I = I + 1

      If CheckBd.GetFieldValue("PART_ITEM_ID") = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & uctlPartLookup.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If

   Next CheckBd
   
   Dim BD As CPackageDetail
   If ShowMode = SHOW_ADD Then
      Set BD = New CPackageDetail
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   Call BD.SetFieldValue("PART_ITEM_ID", uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)))
   Call BD.SetFieldValue("PART_NO", uctlPartLookup.MyTextBox.Text)
   Call BD.SetFieldValue("PART_DESC", uctlPartLookup.MyCombo.Text)
   
   Call BD.SetFieldValue("PART_ITEM_COST", Val(txtPartCost.Text))
   
   Call BD.SetFieldValue("PRO_FROM_DATE", uctlFromProDate.ShowDate)
   Call BD.SetFieldValue("PRO_TO_DATE", uctlToProDate.ShowDate)
   Call BD.SetFieldValue("PRO_ITEM_COST", Val(txtProCost.Text))
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadStockCode(uctlPartLookup.MyCombo, m_Products)
      Set uctlPartLookup.MyCollection = m_Products

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
   
   Set m_Products = New Collection
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_PackageItem = New CPackageDetail
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PackageItem = Nothing
   Set m_Products = Nothing
   
End Sub

Private Sub txtPartCost_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlToProDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlfromProDate_HasChange()
   m_HasModify = True
End Sub
Private Sub txtProCost_Change()
   m_HasModify = True
End Sub
