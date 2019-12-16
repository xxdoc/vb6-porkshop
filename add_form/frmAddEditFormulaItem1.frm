VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditFormulaItem1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditFormulaItem1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   4471
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextBox txtTxAmount 
         Height          =   435
         Left            =   2760
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   2760
         TabIndex        =   0
         Top             =   180
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4965
         TabIndex        =   12
         Top             =   1080
         Width           =   1965
      End
      Begin Threed.SSCommand cmdUnit 
         Height          =   435
         Left            =   4320
         TabIndex        =   3
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaItem1.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblLocationLookup 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPartLookup 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2325
         TabIndex        =   4
         Top             =   1710
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaItem1.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3975
         TabIndex        =   5
         Top             =   1710
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditFormulaItem1.frx":0EFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5625
         TabIndex        =   6
         Top             =   1710
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPartCost 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   1200
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditFormulaItem1"
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

Private m_PackageItem As CFormulaItem

Public ParentIndexTab As Long

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

Private m_Products As Collection
Private m_Locations As Collection
Private m_Sc As CStockCode
'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String
'--------------------------------------------------
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
   
   Call InitNormalLabel(lblPartLookup, MapText("สินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblPartCost, MapText("จำนวน/สูตร"))
   Call InitNormalLabel(lblLocationLookup, MapText("คลัง"))
   Call InitNormalLabel(lblUnit, MapText(""))
   
   uctlPartLookup.MyTextBox.SetKeySearch ("STOCK_NO")
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdUnit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
   Call InitMainButton(cmdUnit, MapText("U"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CFormulaItem
         
         Set BD = TempCollection.Item(ID)
         
         uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, BD.PART_ITEM_ID)
         uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, BD.LOCATION_ID)
         
          txtTxAmount.Text = MyDiffEx(BD.TX_AMOUNT, BD.UNIT_MULTIPLE)
          
         UnitID = BD.UNIT_TRAN_ID
         Multiple = BD.UNIT_MULTIPLE
         UnitName = BD.UNIT_TRAN_NAME
         
         Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
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
         
         Call ParentForm.RefreshGrid(ParentIndexTab)
         Exit Sub
      End If
      
      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
       uctlPartLookup.MyCombo.ListIndex = -1
       uctlLocationLookup.MyCombo.ListIndex = -1
      txtTxAmount.Text = ""
   End If
   Call QueryData(True)
   
   Call uctlPartLookup.SetFocus
   
   Call ParentForm.RefreshGrid(ParentIndexTab)
   
   m_HasModify = True
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
      
   If Not VerifyCombo(lblLocationLookup, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim CheckBd As CFormulaItem
   For Each CheckBd In TempCollection
      I = I + 1

      If CheckBd.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex)) And CheckBd.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & uctlPartLookup.MyCombo.Text & "ในคลัง " & uctlLocationLookup.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   
   Next CheckBd
   
   Dim BD As CFormulaItem
   If ShowMode = SHOW_ADD Then
      Set BD = New CFormulaItem
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   BD.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   BD.PART_NO = uctlPartLookup.MyTextBox.Text
   BD.PART_DESC = uctlPartLookup.MyCombo.Text
   
   BD.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
   BD.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
   
   BD.UNIT_TRAN_ID = UnitID
   BD.UNIT_MULTIPLE = Multiple
   BD.UNIT_TRAN_NAME = UnitName
   
   BD.UNIT_NAME = UnitMName
   
   BD.TX_AMOUNT = Val(txtTxAmount.Text) * Multiple
      
   SaveData = True
End Function
Private Sub cmdUnit_Click()
   frmChangeUnit.HeaderText = MapText("เปลี่ยนหน่วย")
   frmChangeUnit.UnitID = UnitID
   frmChangeUnit.Multiple = Multiple
   frmChangeUnit.UnitName = UnitName
   frmChangeUnit.UnitMName = UnitMName
   
   Load frmChangeUnit
   frmChangeUnit.Show 1
   
   UnitID = frmChangeUnit.UnitID
   Multiple = frmChangeUnit.Multiple
   UnitName = frmChangeUnit.UnitName
   UnitMName = frmChangeUnit.UnitMName
   
   Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   
   Unload frmChangeUnit
   Set frmChangeUnit = Nothing
   m_HasModify = True
   
   Call cmdNext.SetFocus

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadStockCode(uctlPartLookup.MyCombo, m_Products)
      Set uctlPartLookup.MyCollection = m_Products
      
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
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
   Set m_PackageItem = New CFormulaItem
   Set m_Locations = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_PackageItem = Nothing
   Set m_Products = Nothing
   Set m_Locations = Nothing
End Sub

Private Sub txtTxAmount_Change()
   m_HasModify = True
End Sub

Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPartLookup_Change()
On Error Resume Next
Dim ID As Long

   ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If ID > 0 Then
      Set m_Sc = GetObject("CStockCode", m_Products, Trim(Str(ID)))
      
      UnitID = m_Sc.UNIT_ID
      Multiple = m_Sc.UNIT_AMOUNT
      UnitName = m_Sc.UNIT_NAME
      UnitMName = m_Sc.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   Else
      lblUnit.Caption = ""
   End If
   
   
   m_HasModify = True

End Sub

