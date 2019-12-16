VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmJournal 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
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
      Begin Xivess.uctlDate uctlJournalDate 
         Height          =   405
         Left            =   5940
         TabIndex        =   1
         Top             =   960
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboDepartMent 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1860
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   2955
      End
      Begin Xivess.uctlTextBox txtApArName 
         Height          =   435
         Left            =   1560
         TabIndex        =   2
         Top             =   1410
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   767
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
         Height          =   4725
         Left            =   180
         TabIndex        =   9
         Top             =   3000
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8334
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
         Column(1)       =   "frmJournal.frx":27A2
         Column(2)       =   "frmJournal.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmJournal.frx":290E
         FormatStyle(2)  =   "frmJournal.frx":2A6A
         FormatStyle(3)  =   "frmJournal.frx":2B1A
         FormatStyle(4)  =   "frmJournal.frx":2BCE
         FormatStyle(5)  =   "frmJournal.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmJournal.frx":2D5E
      End
      Begin Xivess.uctlTextBox txtJournalCode 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   405
         Left            =   5970
         TabIndex        =   4
         Top             =   1860
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJournalDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4440
         TabIndex        =   22
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label lblJournalCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   19
         Top             =   2340
         Width           =   1185
      End
      Begin VB.Label lblApArName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   2340
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   7
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmJournal.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   8
         Top             =   1500
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
         MouseIcon       =   "frmJournal.frx":3250
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
         MouseIcon       =   "frmJournal.frx":356A
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
         MouseIcon       =   "frmJournal.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Journal As CJournal
Private m_TempJournal As CJournal
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Private m_MasterRef As CMasterRef
Public OKClick As Boolean

Public HeaderText As String
Public ApArInd As Long
Private ApArText As String

Public JournalType As Long

Private Sub cmdPasswd_Click()

End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   If Not VerifyAccessRight("GL_JOURNAL_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmAddEditJournal.JournalType = JournalType
   frmAddEditJournal.HeaderText = MapText("เพิ่มข้อมูลสมุดรายวัน")
   frmAddEditJournal.ShowMode = SHOW_ADD
   Load frmAddEditJournal
   frmAddEditJournal.Show 1
   
   OKClick = frmAddEditJournal.OKClick
   
   Unload frmAddEditJournal
   Set frmAddEditJournal = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtApArName.Text = ""
   txtJournalCode.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
   cboDepartment.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyAccessRight("GL_JOURNAL_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   Call m_Journal.SetFieldValue("JOURNAL_ID", ID)
   If Not glbDaily.DeleteJournal(m_Journal, IsOK, True, glbErrorLog) Then
      Call m_Journal.SetFieldValue("JOURNAL_ID", -1)
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

   If Not VerifyAccessRight("GL_JOURNAL_QUERY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   frmAddEditJournal.JournalType = JournalType
   frmAddEditJournal.ID = ID
   frmAddEditJournal.HeaderText = MapText("แก้ไขข้อมูลสมุดรายวัน")
   frmAddEditJournal.ShowMode = SHOW_EDIT
   Load frmAddEditJournal
   frmAddEditJournal.Show 1
   
   OKClick = frmAddEditJournal.OKClick
   
   Unload frmAddEditJournal
   Set frmAddEditJournal = Nothing
               
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
      
      Call LoadMaster(cboDepartment, , , , MASTER_DEPARTMENT)
      
      Call InitJournalOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If Not VerifyAccessRight("GL_JOURNAL_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      Call m_Journal.SetFieldValue("JOURNAL_ID", -1)
      Call m_Journal.SetFieldValue("FROM_DATE", uctlJournalDate.ShowDate)
      Call m_Journal.SetFieldValue("TO_DATE", uctlJournalDate.ShowDate)
      Call m_Journal.SetFieldValue("JOURNAL_TYPE", JournalType)
      Call m_Journal.SetFieldValue("JOURNAL_NO", txtJournalCode.Text)
      Call m_Journal.SetFieldValue("APAR_NAME", txtApArName.Text)
      Call m_Journal.SetFieldValue("POST_FLAG", Check2Flag(chkPostFlag.Value))
      Call m_Journal.SetFieldValue("DEPARTMENT_ID", cboDepartment.ItemData(Minus2Zero(cboDepartment.ListIndex)))
      Call m_Journal.SetFieldValue("ORDER_BY", cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex)))
      Call m_Journal.SetFieldValue("ORDER_TYPE", cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex)))
      If Not glbDaily.QueryJournal(m_Journal, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
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
   Col.Width = 2460
   Col.Caption = MapText("เลขที่เอกสาร")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2850
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6165
   Col.Caption = MapText("ชื่อลูกค้า/ผู้ค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 0
   Col.Visible = False
   Col.Caption = MapText("POST_FLAG")
   
   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
      
   Call InitGrid
   
   Call InitNormalLabel(lblApArName, MapText("ชื่อลูกค้า/ผู้ค้า"))
   Call InitNormalLabel(lblDepartment, MapText("แผนก"))
   Call InitNormalLabel(lblJournalDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblJournalCode, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call txtApArName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboDepartment)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Call InitCheckBox(chkPostFlag, "POST แล้ว")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
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
   m_TableName = "USER_GROUP"
   
   Set m_Journal = New CJournal
   Set m_TempJournal = New CJournal
   Set m_Rs = New ADODB.Recordset

   Set m_MasterRef = New CMasterRef
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_MasterRef = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(5)
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
   Call m_TempJournal.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempJournal.GetFieldValue("JOURNAL_ID")
   Values(2) = m_TempJournal.GetFieldValue("JOURNAL_NO")
   Values(3) = DateToStringExtEx2(m_TempJournal.GetFieldValue("JOURNAL_DATE"))
   Values(4) = m_TempJournal.GetFieldValue("APAR_NAME")
   Values(5) = m_TempJournal.GetFieldValue("POST_FLAG")
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

