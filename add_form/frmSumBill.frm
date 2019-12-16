VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmSumBill 
   BackColor       =   &H80000000&
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
   Icon            =   "frmSumBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   13860
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10575
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   18653
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   7635
         Left            =   60
         TabIndex        =   3
         Top             =   2130
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   13467
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
         Column(1)       =   "frmSumBill.frx":27A2
         Column(2)       =   "frmSumBill.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmSumBill.frx":290E
         FormatStyle(2)  =   "frmSumBill.frx":2A6A
         FormatStyle(3)  =   "frmSumBill.frx":2B1A
         FormatStyle(4)  =   "frmSumBill.frx":2BCE
         FormatStyle(5)  =   "frmSumBill.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmSumBill.frx":2D5E
      End
      Begin Rebate.uctlTextBox txtSumBillDesc 
         Height          =   405
         Left            =   2100
         TabIndex        =   0
         Top             =   960
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   714
      End
      Begin Rebate.uctlDate uctlFromDate 
         Height          =   405
         Left            =   6420
         TabIndex        =   12
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Rebate.uctlDate uctlToDate 
         Height          =   405
         Left            =   6420
         TabIndex        =   14
         Top             =   1380
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Rebate.uctlTextBox txtReceiptDoNo 
         Height          =   405
         Left            =   2100
         TabIndex        =   16
         Top             =   1380
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   714
      End
      Begin VB.Label lblReceiptDoNo 
         Alignment       =   1  'Right Justify
         Caption         =   "lblReceiptDoNo"
         Height          =   435
         Left            =   270
         TabIndex        =   17
         Top             =   1560
         Width           =   1755
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5160
         TabIndex        =   15
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5160
         TabIndex        =   13
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label lblSumBillDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   270
         TabIndex        =   11
         Top             =   990
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11790
         TabIndex        =   1
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSumBill.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   11790
         TabIndex        =   2
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3300
         TabIndex        =   6
         Top             =   9870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSumBill.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   30
         TabIndex        =   4
         Top             =   9870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSumBill.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1650
         TabIndex        =   5
         Top             =   9870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12135
         TabIndex        =   8
         Top             =   9870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10485
         TabIndex        =   7
         Top             =   9870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSumBill.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmSumBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_SumBill As CSumBill
Private m_TempSumBill As CSumBill
Private m_Rs As ADODB.Recordset

Public OKClick As Boolean

Public DocumentType As Long
Public HeaderText As String
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim TempStr As String
Dim Programowner As String

   Programowner = glbParameterObj.Programowner
   
   frmAddEditSumBill.HeaderText = MapText("เพิ่มข้อมูล" & SellDoctype2Text(DocumentType))
   frmAddEditSumBill.ShowMode = SHOW_ADD
   Load frmAddEditSumBill
   frmAddEditSumBill.Show 1
   
   OKClick = frmAddEditSumBill.OKClick

   Unload frmAddEditSumBill
   Set frmAddEditSumBill = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtSumBillDesc.Text = ""
   txtReceiptDoNo.Text = ""
   
   uctlFromDate.ShowDate = -1
   uctlToDate.ShowDate = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   If Not VerifyLockDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Sub
   Else
      If Not VerifyLockReceiptDate(InternalDateToDateExGrid(GridEX1.Value(3)), InternalDateToDateExGrid(GridEX1.Value(3))) Then
         glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
   
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   m_SumBill.SUM_BILL_ID = ID
   If Not glbDaily.DeleteSumBill(m_SumBill, IsOK, True, glbErrorLog) Then
      m_SumBill.SUM_BILL_ID = -1
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

   Dim Programowner As String
   Programowner = glbParameterObj.Programowner
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   
   frmAddEditSumBill.ID = ID
   frmAddEditSumBill.DocumentType = DocumentType
   frmAddEditSumBill.HeaderText = MapText("แก้ไขข้อมูล" & SellDoctype2Text(DocumentType))
   frmAddEditSumBill.ShowMode = SHOW_EDIT
   Load frmAddEditSumBill
   frmAddEditSumBill.Show 1
   
   OKClick = frmAddEditSumBill.OKClick
   
   Unload frmAddEditSumBill
   Set frmAddEditSumBill = Nothing
   
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
      
      uctlFromDate.ShowDate = Now
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
      
      m_SumBill.SUM_BILL_ID = -1
      m_SumBill.SUM_BILL_DESC = PatchWildCard(txtSumBillDesc.Text)
      m_SumBill.RECEIPT_DOC_NO = PatchWildCard(txtReceiptDoNo.Text)
      m_SumBill.FROM_DATE = uctlFromDate.ShowDate
      m_SumBill.TO_DATE = uctlToDate.ShowDate
      
      If Not glbDaily.QuerySumBill(m_SumBill, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 4000
   Col.Caption = MapText("รายละเอียด")
      
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("วันที่เอกสาร")
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1450
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวนเงิน")
   
   Set Col = Nothing
   Set fmsTemp = Nothing
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   
   Call InitGrid
   
   Call InitNormalLabel(lblSumBillDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblReceiptDoNo, MapText("หมายเลขใบวางบิล"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
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
   
   MasterInd = "1"
   Set m_SumBill = New CSumBill
   Set m_TempSumBill = New CSumBill
   Set m_Rs = New ADODB.Recordset
   
   Call InitFormLayout
   Call EnableForm(Me, True)
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
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_SumBill = Nothing
   Set m_TempSumBill = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I  As Byte
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
   Call m_TempSumBill.PopulateFromRS(1, m_Rs)
   
   I = 0
   I = I + 1
   Values(I) = m_TempSumBill.SUM_BILL_ID
   I = I + 1
   Values(I) = m_TempSumBill.SUM_BILL_DESC
   I = I + 1
   Values(I) = DateToStringExtEx2(m_TempSumBill.DOCUMENT_DATE)
   I = I + 1
   Values(I) = FormatNumber(m_TempSumBill.PAID_AMOUNT - m_TempSumBill.CREDIT_AMOUNT + m_TempSumBill.DEBIT_AMOUNT)
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
