VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmKeyAccount 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmKeyAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   6800
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin GridEX20.GridEX GridEX1 
         Height          =   3165
         Left            =   60
         TabIndex        =   0
         Top             =   30
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   5583
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
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmKeyAccount.frx":27A2
         Column(2)       =   "frmKeyAccount.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmKeyAccount.frx":290E
         FormatStyle(2)  =   "frmKeyAccount.frx":2A6A
         FormatStyle(3)  =   "frmKeyAccount.frx":2B1A
         FormatStyle(4)  =   "frmKeyAccount.frx":2BCE
         FormatStyle(5)  =   "frmKeyAccount.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmKeyAccount.frx":2D5E
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3300
         TabIndex        =   3
         Top             =   3270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmKeyAccount.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   30
         TabIndex        =   1
         Top             =   3270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmKeyAccount.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1650
         TabIndex        =   2
         Top             =   3270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   10185
         TabIndex        =   4
         Top             =   3270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   4695
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   8281
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLocationSale 
         Height          =   315
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3480
         Width           =   1755
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   2325
         Left            =   60
         TabIndex        =   7
         Top             =   1080
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   4101
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
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmKeyAccount.frx":356A
         Column(2)       =   "frmKeyAccount.frx":3632
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmKeyAccount.frx":36D6
         FormatStyle(2)  =   "frmKeyAccount.frx":3832
         FormatStyle(3)  =   "frmKeyAccount.frx":38E2
         FormatStyle(4)  =   "frmKeyAccount.frx":3996
         FormatStyle(5)  =   "frmKeyAccount.frx":3A6E
         ImageCount      =   0
         PrinterProperties=   "frmKeyAccount.frx":3B26
      End
      Begin Rebate.uctlTextLookup uctlCustomer 
         Height          =   435
         Left            =   1800
         TabIndex        =   8
         Top             =   3480
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtDesc 
         Height          =   435
         Left            =   1875
         TabIndex        =   6
         Top             =   600
         Width           =   9885
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlSale 
         Height          =   435
         Left            =   1875
         TabIndex        =   5
         Top             =   120
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   767
      End
      Begin VB.Label lblLocationSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8640
         TabIndex        =   20
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Top             =   3600
         Width           =   975
      End
      Begin Threed.SSCommand cmdOk2 
         Height          =   525
         Left            =   8535
         TabIndex        =   13
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmKeyAccount.frx":3CFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit2 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10185
         TabIndex        =   14
         Top             =   4080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit2 
         Height          =   525
         Left            =   1650
         TabIndex        =   11
         Top             =   4080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd2 
         Height          =   525
         Left            =   30
         TabIndex        =   10
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmKeyAccount.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete2 
         Height          =   525
         Left            =   3300
         TabIndex        =   12
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmKeyAccount.frx":4332
         ButtonStyle     =   3
      End
      Begin VB.Label lblSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmKeyAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SearchMode As Boolean
Private m_HasActivate As Boolean
Private m_KeyAccount As CKeyAccount
Private m_TempKeyAccount As CKeyAccount
Private m_Rs As ADODB.Recordset

Private m_HasModify As Boolean

Public OKClick As Boolean
Public HeaderText As String
Private ID As Long
Private IDD As Long

Public ShowMode As SHOW_MODE_TYPE
Public ShowModeChild As SHOW_MODE_TYPE
Private CustomerColls As Collection
Private Sub cboLocationSale_Click()
   m_HasModify = True
End Sub
Private Sub cboLocationSale_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim TempStr As String
   
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   ShowMode = SHOW_ADD
   Call ChangeSearchMode(False)
   Call SetNewForAdd
   Call uctlSale.MyTextBox.SetFocus
   Call RefreshGrid
   
   m_HasModify = False
End Sub
Private Sub cmdAdd2_Click()
   ShowModeChild = SHOW_ADD
   
   '----------------------------- Check For Prevent KeyPress ------------------------'
   If Not cmdAdd2.Enabled Then
      Exit Sub
   End If
   '----------------------------- Check For Prevent KeyPress ------------------------'
   
   '----------------------------- Save Item And Refresh ------------------------'
   Call SaveDataItem
   Call RefreshGrid
   '----------------------------- Save Item And Refresh ------------------------'
   
   '------------------------------  Clear Item -------------------------------------'
   uctlCustomer.MyCombo.ListIndex = -1
   uctlCustomer.MyTextBox.Text = ""
   cboLocationSale.ListIndex = -1
   '------------------------------  Clear Item -------------------------------------'
   
   '------------------------------ Set FoGus on First -----------------------------'
   Call uctlCustomer.MyTextBox.SetFocus
   '------------------------------ Set FoGus on First -----------------------------'
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim ID1 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID1 = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_KeyAccount.KEY_ACCOUNT_ID = ID1
   If Not glbDaily.DeleteKeyAccount(m_KeyAccount, IsOK, True, glbErrorLog) Then
      m_KeyAccount.KEY_ACCOUNT_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdDelete2_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete2.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX2.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX2.Value(2)
   ID1 = GridEX2.Value(1)
   
   If ID1 <= 0 Then
      m_KeyAccount.KeyAccountDetail.Remove (ID2)
   Else
      m_KeyAccount.KeyAccountDetail.Item(ID2).Flag = "D"
   End If
   
   Call RefreshGrid
   m_HasModify = True
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim OKClick As Boolean
Dim TempStr As String
   
   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   ID = Val(GridEX1.Value(1))
   
   Call ChangeSearchMode(False)
   
   m_KeyAccount.QueryFlag = 1
   Call QueryData
   
   ShowMode = SHOW_EDIT
   
   Call uctlSale.MyTextBox.SetFocus
   m_HasModify = False
End Sub

Private Sub cmdEdit2_Click()
Dim NewID As Long
   
   If Not cmdEdit2.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   
   Call EnEditMode
   ShowModeChild = SHOW_EDIT
   If cmdEdit2.Tag = 1 Then
      IDD = Val(GridEX2.Value(2))
      Call QueryDataItem
      
      cmdEdit2.Tag = 2     ' UPDATE MODE
      Call InitMainButton(cmdEdit2, MapText("บันทึก (F3)"))
   ElseIf cmdEdit2.Tag = 2 Then
      Call SaveDataItem
      
      cmdEdit2.Tag = 1     ' EDIT QUERY MODE
      Call InitMainButton(cmdEdit2, MapText("แก้ไข (F3)"))
      
      Call RefreshGrid
      
      NewID = GetNextID(IDD, m_KeyAccount.KeyAccountDetail)
      If IDD = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call RefreshGrid
         Exit Sub
      End If
      
      IDD = NewID
      Call QueryDataItem
      
      cmdEdit2.Tag = 2     ' UPDATE MODE
      Call InitMainButton(cmdEdit2, MapText("บันทึก (F3)"))
      
      uctlCustomer.SetFocus
   End If
End Sub

Private Sub cmdExit2_Click()
      
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   Call EnAllMode
   Call ClearAllMode
   Call ChangeSearchMode(True)
   
   Call SetNewForAdd
   Call QueryData
End Sub

Private Sub cmdOk2_Click()
Dim oMenu As CPopupMenu
Dim lMenuChosen  As Long

   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออกจากหน้าจอ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      ID = m_KeyAccount.KEY_ACCOUNT_ID
      m_KeyAccount.QueryFlag = 1
      QueryData
      m_HasModify = False
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      OKClick = True
      Unload Me
   End If

End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadEmployee(Nothing, uctlSale.MyCombo)
      Set uctlSale.MyCollection = m_EmployeeColl
      
      m_KeyAccount.QueryFlag = -1
      
      Call InitGrid
      Call InitGrid2
      Call QueryData
   End If
End Sub

Private Sub QueryData()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
      
   Call EnableForm(Me, False)
   
   m_KeyAccount.KEY_ACCOUNT_ID = ID
   If Not glbDaily.QueryKeyAccount(m_KeyAccount, m_Rs, ItemCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call InitGrid
   
   If ItemCount = 1 Then
      Call m_KeyAccount.PopulateFromRS(1, m_Rs)
      
      uctlSale.MyCombo.ListIndex = IDToListIndex(uctlSale.MyCombo, m_KeyAccount.SALE_ID)
      txtDesc.Text = m_KeyAccount.KEY_ACCOUNT_DESC
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call RefreshGrid
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      If SearchMode Then
         Call cmdAdd_Click
      Else
         Call cmdAdd2_Click
      End If
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      If SearchMode Then
         Call cmdDelete_Click
      Else
         Call cmdDelete2_Click
      End If
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      If SearchMode Then
         Call cmdEdit_Click
      Else
         Call cmdEdit2_Click
      End If
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOk2_Click
      KeyCode = 0
   End If
End Sub
Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 4000
   Col.Caption = MapText("พนักงานขาย")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = ScaleWidth - 4200
   Col.Caption = MapText("รายละเอียด")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX2.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("รหัสลูกค้า")

   Set Col = GridEX2.Columns.add '4
   Col.Width = ScaleWidth - 4900
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX2.Columns.add '5
   Col.Width = 2500
   Col.Caption = MapText("สาขา")
   
End Sub
Private Sub InitFormLayout()
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   SearchMode = True
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitNormalLabel(lblSale, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   
   Call InitNormalLabel(lblCustomer, MapText("ลูกค้า"))
   Call InitNormalLabel(lblLocationSale, MapText("สาขา"))
   
   Call InitCombo(cboLocationSale)
   
   SSFrame2.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   cmdOk2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit2.Tag = 1     ' EDIT QUERY MODE
   
   cmdDelete2.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit2, MapText("กลับ (ESC)"))
   Call InitMainButton(cmdOk2, MapText("บันทึก (F2)"))
   Call InitMainButton(cmdAdd2, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit2, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete2, MapText("ลบ (F6)"))
   
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   m_HasActivate = False
   
   Set m_KeyAccount = New CKeyAccount
   Set m_TempKeyAccount = New CKeyAccount
   Set m_Rs = New ADODB.Recordset
   Set CustomerColls = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_KeyAccount = Nothing
   Set CustomerColls = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_TempKeyAccount.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempKeyAccount.KEY_ACCOUNT_ID
   Values(2) = m_TempKeyAccount.SALE_LONG_NAME & " " & m_TempKeyAccount.SALE_LAST_NAME
   Values(3) = m_TempKeyAccount.KEY_ACCOUNT_DESC
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   If SearchMode Then
      SSFrame1.Visible = True
      SSFrame1.Enabled = True
      SSFrame2.Visible = False
      SSFrame2.Enabled = False
      
      SSFrame1.Width = ScaleWidth
      SSFrame1.Height = ScaleHeight
      SSFrame1.Top = 0
      GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
      GridEX1.Height = SSFrame1.Height - GridEX1.Top - 650
      cmdExit.Left = ScaleWidth - cmdExit.Width - 100
      
      cmdAdd.Top = SSFrame1.Height - 580
      cmdEdit.Top = SSFrame1.Height - 580
      cmdDelete.Top = SSFrame1.Height - 580
      cmdExit.Top = SSFrame1.Height - 580
      cmdExit.Left = ScaleWidth - cmdExit.Width - 100
      
      cmdAdd.Enabled = True
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      cmdExit.Enabled = True
            
   Else
      SSFrame1.Visible = False
      SSFrame1.Enabled = False
      SSFrame2.Visible = True
      SSFrame2.Enabled = True
      
      SSFrame2.Width = ScaleWidth
      SSFrame2.Height = ScaleHeight
      SSFrame2.Top = 0
      
      GridEX2.Width = ScaleWidth - 2 * GridEX2.Left
      GridEX2.Height = SSFrame2.Height - GridEX2.Top - 1120
      
      lblCustomer.Top = SSFrame2.Height - 1080
      lblLocationSale.Top = SSFrame2.Height - 1080
      uctlCustomer.Top = SSFrame2.Height - 1080
      cboLocationSale.Top = SSFrame2.Height - 1080
      txtDesc.Width = SSFrame2.Width - txtDesc.Left - 100
      
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
      cmdExit.Enabled = False
         
      cmdEdit2.Tag = 1
      cmdAdd2.Top = SSFrame2.Height - 580
      cmdEdit2.Top = SSFrame2.Height - 580
      cmdDelete2.Top = SSFrame2.Height - 580
      cmdExit2.Top = SSFrame2.Height - 580
      cmdOk2.Top = SSFrame2.Height - 580
      cmdExit2.Left = ScaleWidth - cmdExit2.Width - 100
      cmdOk2.Left = cmdExit2.Left - cmdOk2.Width - 100
      
   End If
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   End If
End Sub
Private Sub GridEX2_DblClick()
   Call cmdEdit2_Click
End Sub
Private Sub GridEX2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit2_Click
      KeyCode = 0
   End If
End Sub
Private Sub txtDesc_Change()
   m_HasModify = True
End Sub
Private Sub uctlCustomer_Change()
Dim ID1 As Long
Dim ID2 As Long
Static OldID1 As Long
Static OldID2 As Long
Dim AparMas As CAPARMas
   
   ID1 = uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex))
   ID2 = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   
   If (ID1 = OldID1) And (ID2 = OldID2) Then
      Exit Sub
   End If
   If ID1 > 0 Then
      OldID1 = ID1
      OldID2 = ID2
      Set AparMas = m_CustomerColl(Trim(Str(ID1)))
      
      Call LoadMaster(cboLocationSale, , , , MASTER_APARMAS_BRANCH, ID1, , ID2)
      
      If cboLocationSale.ListCount > 0 Then
         Dim BranchID As Long
         Dim Branch As CMasterRef
         BranchID = cboLocationSale.ItemData(Minus2Zero(1))
         
         cboLocationSale.ListIndex = 1
         
      End If
   End If
   
   m_HasModify = True
End Sub
Private Function SaveDataItem() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim I As Long

   If Not VerifyCombo(lblCustomer, uctlCustomer.MyCombo, False) Then
      Exit Function
   End If
      
   If Not VerifyCombo(lblLocationSale, cboLocationSale, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveDataItem = True
      Exit Function
   End If
   
   Dim Kad As CKeyAccountDetail
   For Each Kad In m_KeyAccount.KeyAccountDetail
      I = I + 1
            
      If Kad.CUSTOMER_ID = uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex)) And Kad.BRANCH_ID = cboLocationSale.ItemData(Minus2Zero(cboLocationSale.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & uctlCustomer.MyCombo.Text & "ในสาขา " & cboLocationSale.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If
   Next Kad
   Set Kad = Nothing
   
   If ShowModeChild = SHOW_ADD Then
      Set Kad = New CKeyAccountDetail
      Kad.Flag = "A"
      Call m_KeyAccount.KeyAccountDetail.add(Kad)
   Else
      Set Kad = m_KeyAccount.KeyAccountDetail.Item(IDD)
      If Kad.Flag <> "A" Then
         Kad.Flag = "E"
      End If
   End If
   
   Kad.CUSTOMER_ID = uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex))
   Kad.CUSTOMER_NAME = uctlCustomer.MyCombo.Text
   Kad.CUSTOMER_CODE = uctlCustomer.MyTextBox.Text
   
   Kad.BRANCH_ID = cboLocationSale.ItemData(Minus2Zero(cboLocationSale.ListIndex))
   Kad.BRANCH_NAME = cboLocationSale.Text
   
   SaveDataItem = True
End Function
Public Sub RefreshGrid()
   GridEX2.ItemCount = CountItem(m_KeyAccount.KeyAccountDetail)
   GridEX2.Rebind
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"
   
   If m_KeyAccount.KeyAccountDetail Is Nothing Then
      Exit Sub
   End If
   
   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Dim Kad As CKeyAccountDetail
   If m_KeyAccount.KeyAccountDetail.Count <= 0 Then
      Exit Sub
   End If
   Set Kad = GetItem(m_KeyAccount.KeyAccountDetail, RowIndex, RealIndex)
   If Kad Is Nothing Then
      Exit Sub
   End If

   Values(1) = Kad.KEY_ACCOUNT_DETAIL_ID
   Values(2) = RealIndex
   Values(3) = Kad.CUSTOMER_CODE
   Values(4) = Kad.CUSTOMER_NAME
   Values(5) = Kad.BRANCH_NAME
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyCombo(lblSale, uctlSale.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   If Not CheckUniqueNs(KEY_ACCOUNT, uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex)), ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlSale.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
      
   m_KeyAccount.AddEditMode = ShowMode
   m_KeyAccount.KEY_ACCOUNT_ID = ID
   m_KeyAccount.SALE_ID = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   m_KeyAccount.KEY_ACCOUNT_DESC = txtDesc.Text
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditKeyAccount(m_KeyAccount, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub uctlSale_Change()
Dim ID1 As Long
Dim m_Apm As CMasterRef

Dim SaleID As Long
Static OldSaleID As Long
   
   SaleID = uctlSale.MyCombo.ItemData(Minus2Zero(uctlSale.MyCombo.ListIndex))
   If OldSaleID <> SaleID Then
      OldSaleID = SaleID
   Else
      Exit Sub
   End If
   If SaleID > 0 Then
      uctlCustomer.MyTextBox.Text = ""
      Call LoadCustomerFromLocationSale(uctlCustomer.MyCombo, CustomerColls, , MASTER_APARMAS_BRANCH, SaleID)
      Set uctlCustomer.MyCollection = CustomerColls
   End If
   
   m_HasModify = True
End Sub
Private Sub QueryDataItem()
Dim IsOK As Boolean
Dim ItemCount As Long

   Call EnableForm(Me, False)
      
   If ShowModeChild = SHOW_EDIT Then
      Dim BD As CKeyAccountDetail
         
      Set BD = m_KeyAccount.KeyAccountDetail.Item(IDD)
      
      uctlCustomer.MyCombo.ListIndex = IDToListIndex(uctlCustomer.MyCombo, BD.CUSTOMER_ID)
      cboLocationSale.ListIndex = IDToListIndex(cboLocationSale, BD.BRANCH_ID)
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub EnEditMode()
   cmdAdd2.Enabled = False
   cmdEdit2.Enabled = True
   cmdDelete2.Enabled = False
   
   cmdOk2.Enabled = True
   cmdExit2.Enabled = True
End Sub
Private Sub EnAllMode()
   cmdAdd2.Enabled = True
   cmdEdit2.Enabled = True
   cmdDelete2.Enabled = True
   
   cmdOk2.Enabled = True
   cmdExit2.Enabled = True
End Sub
Private Sub ClearAllMode()
   uctlCustomer.MyCombo.ListIndex = -1
   uctlCustomer.MyTextBox.Text = ""
   cboLocationSale.ListIndex = -1
End Sub
Private Sub ChangeSearchMode(Mode As Boolean)
   SearchMode = Mode
   Call Form_Resize
End Sub
Private Sub SetNewForAdd()
   Set m_KeyAccount = Nothing
   uctlSale.MyCombo.ListIndex = -1
   uctlSale.MyTextBox.Text = ""
   txtDesc.Text = ""
   Set m_KeyAccount = New CKeyAccount
   m_KeyAccount.QueryFlag = -1
   ID = -1
End Sub
