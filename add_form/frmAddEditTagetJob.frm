VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditTagetJob 
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   Icon            =   "frmAddEditTagetJob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   13695
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10440
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   18415
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin GridEX20.GridEX GridEX1 
         Height          =   6645
         Left            =   150
         TabIndex        =   6
         Top             =   3000
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   11721
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
         Column(1)       =   "frmAddEditTagetJob.frx":27A2
         Column(2)       =   "frmAddEditTagetJob.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditTagetJob.frx":290E
         FormatStyle(2)  =   "frmAddEditTagetJob.frx":2A6A
         FormatStyle(3)  =   "frmAddEditTagetJob.frx":2B1A
         FormatStyle(4)  =   "frmAddEditTagetJob.frx":2BCE
         FormatStyle(5)  =   "frmAddEditTagetJob.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditTagetJob.frx":2D5E
      End
      Begin VB.ComboBox CboMonthId 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   870
         Width           =   1395
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   675
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2400
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   1191
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtTagetDesc 
         Height          =   435
         Left            =   1890
         TabIndex        =   4
         Top             =   1800
         Width           =   10965
         _ExtentX        =   8334
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtYearNo 
         Height          =   435
         Left            =   3630
         TabIndex        =   1
         Top             =   840
         Width           =   915
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtInputAmount 
         Height          =   435
         Left            =   8040
         TabIndex        =   3
         Top             =   1320
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1890
         TabIndex        =   2
         Top             =   1320
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lblInputAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   7200
         TabIndex        =   17
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblYearNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   3270
         TabIndex        =   16
         Top             =   930
         Width           =   255
      End
      Begin VB.Label lblMonthId 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   15
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label lblTagetDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   1365
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10395
         TabIndex        =   10
         Top             =   9750
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetJob.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12045
         TabIndex        =   11
         Top             =   9750
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   8
         Top             =   9750
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   7
         Top             =   9750
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetJob.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   9750
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetJob.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditTagetJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_TagetJob As CTagetJob

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Public m_Products As Collection
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_TagetJob.TAGET_JOB_ID = ID
      m_TagetJob.QueryFlag = 1
      If Not glbDaily.QueryTagetJob(m_TagetJob, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_TagetJob.PopulateFromRS(1, m_Rs)
      
      CboMonthId.ListIndex = IDToListIndex(CboMonthId, Val(m_TagetJob.MONTH_ID))
      txtYearNo.Text = m_TagetJob.YEAR_NO + 543
      txtTagetDesc.Text = m_TagetJob.TAGET_JOB_DESC
      uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, m_TagetJob.INPUT_ID)
      txtInputAmount.Text = Val(m_TagetJob.INPUT_AMOUNT)
      
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblYearNo, txtYearNo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblMonthId, CboMonthId, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblInputAmount, txtInputAmount, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   m_TagetJob.AddEditMode = ShowMode
   m_TagetJob.TAGET_JOB_ID = ID
   m_TagetJob.MONTH_ID = CboMonthId.ItemData(Minus2Zero(CboMonthId.ListIndex))
   m_TagetJob.YEAR_NO = Val(txtYearNo.Text) - 543
   m_TagetJob.TAGET_JOB_DESC = txtTagetDesc.Text
   m_TagetJob.INPUT_ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   m_TagetJob.INPUT_AMOUNT = Val(txtInputAmount.Text)
   
   If Not CheckUniqueNs(JOB_TAGET_UNIQUE, Trim((Val(txtYearNo.Text) - 543)), ID, m_TagetJob.MONTH_ID, m_TagetJob.INPUT_ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & uctlProductLookup.MyCombo.Text & " " & CboMonthId.Text & " " & txtYearNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditTagetJob(m_TagetJob, IsOK, True, glbErrorLog) Then
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
Private Sub CboMonthId_Click()
   m_HasModify = True
End Sub
Private Sub CboMonthId_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
   
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditTagetJobDetail.ParentForm = Me
      Set frmAddEditTagetJobDetail.TempCollection = m_TagetJob.TagetJobDetail
      frmAddEditTagetJobDetail.ShowMode = SHOW_ADD
      frmAddEditTagetJobDetail.HeaderText = MapText("เพิ่มเป้าการผลิต")
      Load frmAddEditTagetJobDetail
      frmAddEditTagetJobDetail.Show 1

      OKClick = frmAddEditTagetJobDetail.OKClick

      Unload frmAddEditTagetJobDetail
      Set frmAddEditTagetJobDetail = Nothing
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      Call RefreshGrid
   End If

   If OKClick Then
      m_HasModify = True
   End If
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
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_TagetJob.TagetJobDetail.Remove (ID2)
      Else
         m_TagetJob.TagetJobDetail.Item(ID2).Flag = "D"
      End If
      

   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   
   Call RefreshGrid
   m_HasModify = True

End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditTagetJobDetail.ParentForm = Me
      frmAddEditTagetJobDetail.ID = ID
      Set frmAddEditTagetJobDetail.TempCollection = m_TagetJob.TagetJobDetail
      frmAddEditTagetJobDetail.HeaderText = MapText("แก้ไขเป้าการขายตามสาขาลูกค้า")
      frmAddEditTagetJobDetail.ShowMode = SHOW_EDIT
      Load frmAddEditTagetJobDetail
      frmAddEditTagetJobDetail.Show 1

      OKClick = frmAddEditTagetJobDetail.OKClick

      Unload frmAddEditTagetJobDetail
      Set frmAddEditTagetJobDetail = Nothing
      
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   If OKClick Then
      Call RefreshGrid
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
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
      ID = m_TagetJob.TAGET_JOB_ID
      m_TagetJob.QueryFlag = 1
      QueryData (True)
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
      Me.Refresh
      DoEvents
      
      Call InitThaiMonth(CboMonthId)
      
      Call LoadStockCode(uctlProductLookup.MyCombo, m_Products)
      Set uctlProductLookup.MyCollection = m_Products
      
      Call EnableForm(Me, False)
      
      Call InitGrid1
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_TagetJob.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_TagetJob.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_TagetJob = Nothing
   Set m_Products = Nothing
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
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1500
   Col.Caption = MapText("BATCH_NO")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 2000
   Col.Caption = MapText("ประเภทผลิต")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 2200
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดผลิต")
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblMonthId, MapText("เดือน"))
   Call InitNormalLabel(lblYearNo, MapText("ปี"))
   Call InitNormalLabel(lblTagetDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblProduct, MapText("วัตถุดิบ"))
   Call InitNormalLabel(lblInputAmount, MapText("จำนวน"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call txtYearNo.SetTextLenType(TEXT_STRING, glbSetting.YEAR_TYPE)
   Call txtInputAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitCombo(CboMonthId)
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายละเอียดเป้าผลผลิตที่ได้")
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_TagetJob = New CTagetJob
   Set m_Products = New Collection
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_TagetJob.TagetJobDetail Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CTagetJobDetail
      If m_TagetJob.TagetJobDetail.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_TagetJob.TagetJobDetail, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.TAGET_JOB_DETAIL_ID
      Values(2) = RealIndex
      Values(3) = CR.BATCH_NO
      Values(4) = CR.OUTPUT_DESC
      Values(5) = CR.OUTPUT_AMOUNT
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
  If TabStrip1.SelectedItem.Index = 1 Then
      Call RefreshGrid
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub
Public Sub RefreshGrid()
   
   m_HasModify = True
   
   GridEX1.ItemCount = CountItem(m_TagetJob.TagetJobDetail)
   GridEX1.Rebind
End Sub
Private Sub txtYearNo_Change()
   m_HasModify = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   TabStrip1.Width = GridEX1.Width
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   
End Sub
Private Function CheckUniqueTagetJob(MonthID As Long, YearNo As String, InputID As Long)
Dim Tj As CTagetJob
'Dim TempRs as  ADODB.re
End Function
