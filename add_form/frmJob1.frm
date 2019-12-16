VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmJob1 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmJob1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlDate uctlJobDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   2
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   3825
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   16
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2085
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   3678
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
         Column(1)       =   "frmJob1.frx":27A2
         Column(2)       =   "frmJob1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmJob1.frx":290E
         FormatStyle(2)  =   "frmJob1.frx":2A6A
         FormatStyle(3)  =   "frmJob1.frx":2B1A
         FormatStyle(4)  =   "frmJob1.frx":2BCE
         FormatStyle(5)  =   "frmJob1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmJob1.frx":2D5E
      End
      Begin PorkShop.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1680
         TabIndex        =   1
         Top             =   1320
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlFormula 
         Height          =   435
         Left            =   1680
         TabIndex        =   4
         Top             =   1800
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   767
      End
      Begin VB.Label lblFormula 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   23
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4770
         TabIndex        =   22
         Top             =   1380
         Width           =   1305
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4860
         TabIndex        =   19
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   18
         Top             =   900
         Width           =   1425
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   2400
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJob1.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1530
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJob1.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJob1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   11
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   14
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJob1.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmJob1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Job As CJob
Private m_TempJob As CJob
Private m_Rs As ADODB.Recordset

Private m_Formulas  As Collection

Public OKClick As Boolean
Public HeaderText As String
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim TempStr As String
   
   frmAddEditJob1.HeaderText = MapText("เพิ่มข้อมูลใบงาน")
   frmAddEditJob1.ShowMode = SHOW_ADD
   Load frmAddEditJob1
   frmAddEditJob1.Show 1

   OKClick = frmAddEditJob1.OKClick

   Unload frmAddEditJob1
   Set frmAddEditJob1 = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdClear_Click()
   txtJobNo.Text = ""
   txtPartNo.Text = ""
   uctlJobDate.ShowDate = -1
   uctlToDate.ShowDate = -1
   uctlFormula.MyCombo.ListIndex = IDToListIndex(uctlFormula.MyCombo, -1)
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim ID As Long
Dim InventoryDocID As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   InventoryDocID = GridEX1.Value(2)
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   m_Job.JOB_ID = ID
   m_Job.INVENTORY_DOC_ID = InventoryDocID
   If Not glbDaily.DeleteJob(m_Job, IsOK, True, glbErrorLog) Then
      m_Job.JOB_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim TempStr As String
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   frmAddEditJob1.ID = ID
   frmAddEditJob1.HeaderText = MapText("แก้ไขข้อมูลใบงาน")
   frmAddEditJob1.ShowMode = SHOW_EDIT
   Load frmAddEditJob1
   frmAddEditJob1.Show 1

   OKClick = frmAddEditJob1.OKClick

   Unload frmAddEditJob1
   Set frmAddEditJob1 = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call InitJobOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call LoadFormula(uctlFormula.MyCombo, m_Formulas)
      Set uctlFormula.MyCollection = m_Formulas
      
      uctlJobDate.ShowDate = Now
      uctlToDate.ShowDate = Now
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_Job.JOB_ID = -1
      m_Job.JOB_NO = PatchWildCard(txtJobNo.Text)
      m_Job.STOCK_NO = PatchWildCard(txtPartNo.Text)
      m_Job.FROM_DATE = uctlJobDate.ShowDate
      m_Job.TO_DATE = uctlToDate.ShowDate
      m_Job.FORMULA_ID = uctlFormula.MyCombo.ItemData(Minus2Zero(uctlFormula.MyCombo.ListIndex))
      m_Job.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_Job.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      
      If Not glbDaily.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call InitGrid
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 123 Then
      'Call AddMemoNote
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
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "INVENTORY_DOC_ID"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 2000
   Col.Caption = MapText("หมายเลข JOB")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2300
   Col.Caption = MapText("วันเวลา")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("รหัสสูตร")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = ScaleWidth - 8000
   Col.Caption = MapText("สูตร")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน/สูตร")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblJobDate, MapText("จากวันที่ JOB"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่ JOB"))
   Call InitNormalLabel(lblJobNo, MapText("เลขที่ JOB"))
   Call InitNormalLabel(lblFormula, MapText("สูตร"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   Call InitNormalLabel(lblPartNo, MapText("หมายเลขวัตถุดิบ"))
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Call txtPartNo.SetKeySearch("STOCK_NO")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   m_HasActivate = False
   
   Set m_Job = New CJob
   Set m_TempJob = New CJob
   Set m_Formulas = New Collection
   Set m_Rs = New ADODB.Recordset
      
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Job = Nothing
   Set m_Formulas = Nothing
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
Dim Ji As CJobItem
   
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
   Call m_TempJob.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempJob.JOB_ID
   Values(2) = m_TempJob.INVENTORY_DOC_ID
   Values(3) = m_TempJob.JOB_NO
   Values(4) = DateToStringExtEx3(m_TempJob.JOB_DATE)
   Values(5) = m_TempJob.FORMULA_NO
   Values(6) = m_TempJob.FORMULA_DESC
   Values(7) = FormatNumber(m_TempJob.FORMULA_AMOUNT)
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
