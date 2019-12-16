VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPrintLabel 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditPrintLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7905
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   13944
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTextBox txtItemAmount 
         Height          =   435
         Left            =   8520
         TabIndex        =   3
         Top             =   6600
         Width           =   915
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlBranch 
         Height          =   435
         Left            =   1920
         TabIndex        =   1
         Top             =   6600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextLookup uctlBlock 
         Height          =   435
         Left            =   1920
         TabIndex        =   0
         Top             =   6120
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtMainAmount 
         Height          =   435
         Left            =   2280
         TabIndex        =   10
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4755
         Left            =   150
         TabIndex        =   19
         Top             =   600
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8387
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditPrintLabel.frx":08CA
         Column(2)       =   "frmAddEditPrintLabel.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPrintLabel.frx":0A36
         FormatStyle(2)  =   "frmAddEditPrintLabel.frx":0B92
         FormatStyle(3)  =   "frmAddEditPrintLabel.frx":0C42
         FormatStyle(4)  =   "frmAddEditPrintLabel.frx":0CF6
         FormatStyle(5)  =   "frmAddEditPrintLabel.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPrintLabel.frx":0E86
      End
      Begin Xivess.uctlTextBox txtSumAmount 
         Height          =   435
         Left            =   8775
         TabIndex        =   20
         Top             =   120
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtPackAmount 
         Height          =   435
         Left            =   10725
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6600
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   767
      End
      Begin VB.Label lblPackAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9600
         TabIndex        =   23
         Top             =   6720
         Width           =   1005
      End
      Begin VB.Label lblSumAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6720
         TabIndex        =   22
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblUnitSum 
         Height          =   375
         Left            =   10815
         TabIndex        =   21
         Top             =   120
         Width           =   765
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1740
         TabIndex        =   6
         Top             =   5460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   120
         TabIndex        =   5
         Top             =   5460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPrintLabel.frx":105E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3390
         TabIndex        =   7
         Top             =   5460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPrintLabel.frx":1378
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   120
         Width           =   765
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   5160
         TabIndex        =   4
         Top             =   7170
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblBlock 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   330
         TabIndex        =   17
         Top             =   6240
         Width           =   1485
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   330
         TabIndex        =   16
         Top             =   6720
         Width           =   1485
      End
      Begin VB.Label lblMainAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   225
         TabIndex        =   15
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         Height          =   345
         Left            =   8220
         TabIndex        =   14
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8355
         TabIndex        =   8
         Top             =   5460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   10005
         TabIndex        =   9
         Top             =   5460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7395
         TabIndex        =   13
         Top             =   6720
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmAddEditPrintLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ParentForm As Form
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public TempCollection2 As Collection
Public MainItemAmount As Double
Public AvgPrice As Double
Public UnitAmount As Double
Public UnitPerBasket As Double

Private m_Branchs  As Collection
Private m_Blocks  As Collection
Private m_Mr As CMasterRef

Private Sub cmdAdd_Click()
   ShowMode = SHOW_ADD
   uctlBlock.MyCombo.ListIndex = -1
   uctlBranch.MyCombo.ListIndex = -1
   txtItemAmount.Text = ""
   txtPackAmount.Text = ""
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If

   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)

   If ID1 <= 0 Then
      TempCollection.Remove (ID2)
   Else
      TempCollection.Item(ID2).Flag = "D"
   End If

   Call GetAmount
   GridEX1.ItemCount = CountItem(TempCollection)
   GridEX1.Rebind
   m_HasModify = True

End Sub
Private Sub GetAmount()
Dim II As CPrintLabel
Dim Sum1 As Double

   Sum1 = 0
   
   For Each II In TempCollection
      If II.Flag <> "D" Then
         Sum1 = Sum1 + II.GetFieldValue("ITEM_AMOUNT")
      End If
   Next II
   
   txtSumAmount.Text = Sum1
   
End Sub

Private Sub cmdEdit_Click()
   ShowMode = SHOW_EDIT
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   
   Call QueryData(True)
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
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
      
   Call InitNormalLabel(lblMainAmount, MapText("จำนวนทั้งหมด"))
   Call InitNormalLabel(lblBlock, MapText("บล็อค"))
   Call InitNormalLabel(lblBranch, MapText("สาขา"))
   Call InitNormalLabel(lblItemAmount, MapText("จ.หน่วย"))
   Call InitNormalLabel(lblPackAmount, MapText("จ.ตะกร้า"))
   Call InitNormalLabel(lblSumAmount, MapText("จำนวนในรายการ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วย"))
   Call InitNormalLabel(lblUnitSum, MapText("หน่วย"))
   
   Call txtPackAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtItemAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtMainAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtSumAmount.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMainAmount.Enabled = False
   txtSumAmount.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
   
   Call InitGrid1
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Ivd As CInventoryDoc
Dim iCount As Long
Dim Ei As CLotItem

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Plb As CPrintLabel
         Set Plb = TempCollection.Item(ID)
         
          uctlBlock.MyCombo.ListIndex = IDToListIndex(uctlBlock.MyCombo, Plb.GetFieldValue("BLOCK_ID"))
          uctlBranch.MyCombo.ListIndex = IDToListIndex(uctlBranch.MyCombo, Plb.GetFieldValue("BRANCH_ID"))
          txtItemAmount.Text = Plb.GetFieldValue("ITEM_AMOUNT")
          txtPackAmount.Text = Plb.GetFieldValue("PACK_AMOUNT")
          
      End If
   End If
   
   Call GetAmount
   
   GridEX1.ItemCount = CountItem(TempCollection)
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdNext_Click()
Dim NewID As Long
   
   If ShowMode <> SHOW_EDIT Then
      ShowMode = SHOW_ADD
   End If
   If Not SaveData Then
      Exit Sub
   End If
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
      Else
         ID = NewID
      End If
      
      
   ElseIf ShowMode = SHOW_ADD Then
      uctlBlock.MyCombo.ListIndex = -1
      uctlBranch.MyCombo.ListIndex = -1
      txtItemAmount.Text = ""
      txtPackAmount.Text = ""
   End If
   Call QueryData(True)
   
   Call uctlBlock.MyTextBox.SetFocus
End Sub
Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
'   If Not SaveData Then
'      Exit Sub
'   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim TempID As Long
   
   If Not VerifyCombo(lblBlock, uctlBlock.MyCombo, False) Then
      Exit Function
   End If
      
   If Not VerifyCombo(lblBranch, uctlBranch.MyCombo, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblItemAmount, txtItemAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   TempID = ID
   Dim Plb As CPrintLabel
   For Each Plb In TempCollection
      If ShowMode = SHOW_ADD Then
         If Plb.GetFieldValue("BLOCK_ID") = uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex)) And Plb.GetFieldValue("BRANCH_ID") = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)) Then
            glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลบล็อค ") & Plb.GetFieldValue("BLOCK_NAME") & " และสาขา " & Plb.GetFieldValue("BRANCH_NAME") & " " & MapText("อยู่ในระบบแล้ว")
            glbErrorLog.ShowUserError
            Exit Function
         End If
      Else
         If Plb.GetFieldValue("BLOCK_ID") = uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex)) And Plb.GetFieldValue("BRANCH_ID") = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)) And Not (TempID = ID) Then
            glbErrorLog.LocalErrorMsg = MapText("มีข้อมูลบล็อค ") & Plb.GetFieldValue("BLOCK_NAME") & " และสาขา " & Plb.GetFieldValue("BRANCH_NAME") & " " & MapText("อยู่ในระบบแล้ว")
            glbErrorLog.ShowUserError
         Exit Function
      End If
      End If
   Next
   
   If ShowMode = SHOW_ADD Then
      Set Plb = New CPrintLabel

      Plb.Flag = "A"
      Call TempCollection.add(Plb)
   Else
      Set Plb = TempCollection.Item(ID)
      If Plb.Flag <> "A" Then
         Plb.Flag = "E"
      End If
   End If
   
   Call Plb.SetFieldValue("BLOCK_ID", uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex)))
   Call Plb.SetFieldValue("BLOCK_NAME", uctlBlock.MyCombo.Text)
   Call Plb.SetFieldValue("BRANCH_ID", uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)))
   Call Plb.SetFieldValue("BRANCH_NAME", uctlBranch.MyCombo.Text)
   Call Plb.SetFieldValue("ITEM_AMOUNT", Val(txtItemAmount.Text))
   Call Plb.SetFieldValue("PACK_AMOUNT", Val(txtPackAmount.Text))
   Call Plb.SetFieldValue("TOTAL_PRICE", Val(txtItemAmount.Text) * AvgPrice)
   Call Plb.SetFieldValue("TOTAL_AMOUNT", Val(txtItemAmount.Text) * UnitAmount)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
            
      Call LoadMaster(uctlBlock.MyCombo, m_Branchs, , , MASTER_CUSTOMER_BLOCK)
      Set uctlBlock.MyCollection = m_Branchs
            
      txtMainAmount.Text = MainItemAmount
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
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
   
   Set m_Rs = New ADODB.Recordset
   Set m_Branchs = New Collection
   Set m_Blocks = New Collection
   Set m_Mr = New CMasterRef
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_Branchs = Nothing
   Set m_Blocks = Nothing
   Set m_Mr = Nothing
End Sub
Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.add '3
   Col.Width = 3500
   Col.Caption = MapText("บล็อค")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 3500
   Col.Caption = MapText("สาขา")
   
   Set Col = GridEX1.Columns.add '5
   Col.TextAlignment = jgexAlignRight
   Col.Width = 2000
   Col.Caption = MapText("ตะกร้า")
   
   Set Col = GridEX1.Columns.add '6
   Col.TextAlignment = jgexAlignRight
   Col.Width = 2000
   Col.Caption = MapText("จำนวน")
   
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TempCollection Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim Plb As CPrintLabel
   If TempCollection.Count <= 0 Then
      Exit Sub
   End If
   Set Plb = GetItem(TempCollection, RowIndex, RealIndex)
   If Plb Is Nothing Then
      Exit Sub
   End If

   Values(1) = Plb.GetFieldValue("PRINT_LABEL_ID")
   Values(2) = RealIndex
   Values(3) = Plb.GetFieldValue("BLOCK_NAME")
   Values(4) = Plb.GetFieldValue("BRANCH_NAME")
   Values(5) = FormatNumber(Plb.GetFieldValue("PACK_AMOUNT"))
   Values(6) = FormatNumber(Plb.GetFieldValue("ITEM_AMOUNT"))
      
Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtItemAmount_Change()
   m_HasModify = True
   txtPackAmount.Text = MyDiffEx(Val(txtItemAmount.Text), UnitPerBasket)
End Sub
Private Sub uctlBlock_Change()
Dim ID As Long
   
   ID = uctlBlock.MyCombo.ItemData(Minus2Zero(uctlBlock.MyCombo.ListIndex))
   If ID > 0 Then
      Call LoadMaster(uctlBranch.MyCombo, m_Branchs, , , MASTER_APARMAS_BRANCH, , ID)
      Set uctlBranch.MyCollection = m_Branchs
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlBranch_Change()
   m_HasModify = True
End Sub
