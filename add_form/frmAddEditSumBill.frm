VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddEditSumBill 
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13830
   Icon            =   "frmAddEditSumBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   13830
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1275
      Left            =   1560
      ScaleHeight     =   1215
      ScaleWidth      =   1575
      TabIndex        =   12
      Top             =   -600
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10455
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   18441
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1725
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   979
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   6585
         Left            =   45
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2280
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   11615
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
         Column(1)       =   "frmAddEditSumBill.frx":27A2
         Column(2)       =   "frmAddEditSumBill.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditSumBill.frx":290E
         FormatStyle(2)  =   "frmAddEditSumBill.frx":2A6A
         FormatStyle(3)  =   "frmAddEditSumBill.frx":2B1A
         FormatStyle(4)  =   "frmAddEditSumBill.frx":2BCE
         FormatStyle(5)  =   "frmAddEditSumBill.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditSumBill.frx":2D5E
      End
      Begin Rebate.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   1800
         TabIndex        =   0
         Top             =   750
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   1032
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtSumBillDesc 
         Height          =   435
         Left            =   1800
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   11595
         _ExtentX        =   3307
         _ExtentY        =   767
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   660
         Left            =   0
         TabIndex        =   14
         Top             =   9000
         Visible         =   0   'False
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   1164
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Rebate.uctlTextBox txtPaidAmount 
            Height          =   435
            Left            =   1440
            TabIndex        =   15
            Top             =   120
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin Rebate.uctlTextBox txtDebitAmount 
            Height          =   435
            Left            =   4440
            TabIndex        =   16
            Top             =   120
            Width           =   1875
            _ExtentX        =   2037
            _ExtentY        =   767
         End
         Begin Rebate.uctlTextBox txtCreditAmount 
            Height          =   435
            Left            =   7320
            TabIndex        =   17
            Top             =   120
            Width           =   1875
            _ExtentX        =   2672
            _ExtentY        =   767
         End
         Begin Rebate.uctlTextBox txtAfterDebitCredit 
            Height          =   435
            Left            =   10560
            TabIndex        =   18
            Top             =   120
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   767
         End
         Begin VB.Label lblPaidAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   480
            TabIndex        =   22
            Top             =   195
            Width           =   915
         End
         Begin VB.Label lblCreditAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6480
            TabIndex        =   21
            Top             =   195
            Width           =   705
         End
         Begin VB.Label lblDebitAmount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3480
            TabIndex        =   20
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblAfterDebitCredit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9840
            TabIndex        =   19
            Top             =   195
            Width           =   675
         End
      End
      Begin VB.Label lblSumBillDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1425
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   8850
         TabIndex        =   5
         Top             =   9825
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   810
         Width           =   555
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10515
         TabIndex        =   6
         Top             =   9825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSumBill.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   12165
         TabIndex        =   7
         Top             =   9825
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   30
         TabIndex        =   3
         Top             =   9825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSumBill.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   1620
         TabIndex        =   4
         Top             =   9825
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSumBill.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditSumBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_SumBill As CSumBill

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As SELL_BILLING_DOCTYPE
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_SumBill.SUM_BILL_ID = ID
      
      If Not glbDaily.QuerySumBill(m_SumBill, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_SumBill.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_SumBill.DOCUMENT_DATE
      txtSumBillDesc.Text = m_SumBill.SUM_BILL_DESC
      
      txtPaidAmount.Text = m_SumBill.PAID_AMOUNT
      txtCreditAmount.Text = m_SumBill.CREDIT_AMOUNT
      txtDebitAmount.Text = m_SumBill.DEBIT_AMOUNT
      txtAfterDebitCredit.Text = m_SumBill.PAID_AMOUNT - m_SumBill.CREDIT_AMOUNT + m_SumBill.DEBIT_AMOUNT
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
Dim CBdt As CBillDetail

   If Not (cmdOK.Enabled) Then
      Exit Function
   End If

   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblSumBillDesc, txtSumBillDesc, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   m_SumBill.ShowMode = ShowMode
   m_SumBill.SUM_BILL_ID = ID
   m_SumBill.DOCUMENT_DATE = uctlDocumentDate.ShowDate
   m_SumBill.SUM_BILL_DESC = txtSumBillDesc.Text
   
   m_SumBill.PAID_AMOUNT = Val(txtPaidAmount.Text)
   m_SumBill.CREDIT_AMOUNT = Val(txtCreditAmount.Text)
   m_SumBill.DEBIT_AMOUNT = Val(txtDebitAmount.Text)
   
   Call EnableForm(Me, False)
   
   Call glbDaily.StartTransaction
   
   If Not glbDaily.AddEditSumBill(m_SumBill, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Call glbDaily.RollbackTransaction
      Exit Function
   End If

   Call glbDaily.CommitTransaction

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
   
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
         
   Set frmAddBillDetail.TempCollection = m_SumBill.RcpCnDnItems
   frmAddBillDetail.ShowMode = SHOW_ADD
   frmAddBillDetail.HeaderText = MapText("เพิ่มรายการ")
   
   Load frmAddBillDetail
   frmAddBillDetail.Show 1
   
   OKClick = frmAddBillDetail.OKClick
   
   Unload frmAddBillDetail
   Set frmAddBillDetail = Nothing

   If OKClick Then
      Call GetTotalPriceReceipt

      GridEX1.ItemCount = CountItem(m_SumBill.RcpCnDnItems)
      GridEX1.Rebind
      m_HasModify = True
   End If
   
   Set oMenu = Nothing
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
      m_SumBill.RcpCnDnItems.Remove (ID2)
   Else
      m_SumBill.RcpCnDnItems.Item(ID2).Flag = "D"
   End If

   Call GetTotalPriceReceipt
   GridEX1.ItemCount = CountItem(m_SumBill.RcpCnDnItems)
   GridEX1.Rebind
   m_HasModify = True
   
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
      ID = m_SumBill.SUM_BILL_ID
      m_SumBill.QueryFlag = 1
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
Private Sub cmdPrint_Click()
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu
Dim OKClick As Boolean
Dim iCount As Long
Dim ReportFlag As Boolean
Dim ReportKey As String
Dim Report As CReportInterface
Dim Rc As CReportConfig
Dim EditMode As SHOW_MODE_TYPE
Dim ReportMode As Long
   
   
   If m_HasModify Or (m_SumBill.SUM_BILL_ID <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("พิมพ์ใบสรุปยอดวางบิล", "-", "ปรับค่าหน้ากระดาษ")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   ReportMode = 1
   ReportFlag = False
   
   If lMenuChosen = 1 Then
      ReportKey = "CReportNormalSumBill001"

      Set Report = New CReportNormalSumBill001

      If Val(txtDebitAmount.Text) > 0 Then
         Call Report.AddParam("Y", "HAVE_DN")
      Else
         Call Report.AddParam("N", "HAVE_DN")
      End If
      ReportFlag = True
   ElseIf lMenuChosen = 3 Then
      ReportMode = 1
      ReportKey = "CReportNormalSumBill001"

      Set Rc = New CReportConfig
      Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
      Call Rc.QueryData(1, m_Rs, iCount)

      HeaderText = MapText("ใบสรุปยอดวางบิล")

      If Not m_Rs.EOF Then
         Call Rc.PopulateFromRS(1, m_Rs)
         EditMode = SHOW_EDIT
      Else
         EditMode = SHOW_ADD
      End If
   Else
      Exit Sub
   End If

   If Not Report Is Nothing Then
      Call Report.AddParam(lMenuChosen, "REPORT_TYPE")
      Call Report.AddParam(m_SumBill.SUM_BILL_ID, "SUM_BILL_ID")
      Call Report.AddParam(m_SumBill.DOCUMENT_DATE, "DOCUMENT_DATE")
      Call Report.AddParam(m_SumBill.SUM_BILL_DESC, "SUM_BILL_DESC")
      Call Report.AddParam(ReportKey, "REPORT_KEY")
   End If

   Call EnableForm(Me, False)
   If ReportFlag Then
      frmReport.ClassName = ReportKey
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = ""
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
      Set Report = Nothing
   Else
      frmReportConfig.ReportMode = ReportMode
      frmReportConfig.ShowMode = EditMode
      frmReportConfig.ID = Rc.GetFieldValue("REPORT_CONFIG_ID")
      frmReportConfig.ReportKey = ReportKey
      frmReportConfig.HeaderText = HeaderText
      Load frmReportConfig
      frmReportConfig.Show 1

      Unload frmReportConfig
      Set frmReportConfig = Nothing
   End If
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call EnableForm(Me, False)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_SumBill.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_SumBill.QueryFlag = 0
         Call QueryData(False)
      End If

      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long

   If InUsed = 1 Then
      Exit Sub
   End If

   InUsed = 1

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
      'Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdPrint_Click
      KeyCode = 0
   End If

   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_SumBill = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
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
   Col.Width = 2220
   Col.Caption = MapText("เลขที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2730
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 3000
   Col.Caption = MapText("ชื่อลูกค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ยอดบิล")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ลดหนี้")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("เพิ่มหนี้")

   Set Col = GridEX1.Columns.add '7
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("รวม")
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame4.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Call InitNormalLabel(lblSumBillDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่"))
   
   Call InitNormalLabel(lblPaidAmount, MapText("ยอดชำระ"))
   Call InitNormalLabel(lblDebitAmount, MapText("เพิ่มหนี้"))
   Call InitNormalLabel(lblCreditAmount, MapText("ลดหนี้"))
   Call InitNormalLabel(lblAfterDebitCredit, MapText("รวม"))
   
   txtPaidAmount.Enabled = False
   txtDebitAmount.Enabled = False
   txtCreditAmount.Enabled = False
   txtAfterDebitCredit.Enabled = False
   
   
   Call txtSumBillDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPaidAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtDebitAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtCreditAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtAfterDebitCredit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   GridEX1.Visible = True

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์ (F10)"))
   
   Call InitGrid1

   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16

   Dim T As Object
   TabStrip1.Tabs.Clear

   SSFrame4.Visible = True
   
   Set T = TabStrip1.Tabs.add()
   T.Caption = MapText("รายการใบวางบิล")
   T.Tag = DocumentType & "-Bdt"
   
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
   Set m_SumBill = New CSumBill
   
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   
   If m_SumBill.RcpCnDnItems Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim Rc As CBillDetail
   If m_SumBill.RcpCnDnItems.Count <= 0 Then
      Exit Sub
   End If
   Set Rc = GetItem(m_SumBill.RcpCnDnItems, RowIndex, RealIndex)
   If Rc Is Nothing Then
      Exit Sub
   End If

   Values(1) = Rc.BILL_DETAIL_ID
   Values(2) = RealIndex
   Values(3) = Rc.SUMMARY_DOC_NO
   Values(4) = DateToStringExtEx2(Rc.SUMMARY_DOC_DATE)
   Values(5) = Rc.APAR_CODE
   Values(6) = Rc.APAR_NAME
   Values(7) = FormatNumber(Rc.PAID_AMOUNT)
   Values(8) = FormatNumber(Rc.CREDIT_AMOUNT)
   Values(9) = FormatNumber(Rc.DEBIT_AMOUNT)
   Values(10) = FormatNumber(Rc.PAID_AMOUNT - Rc.CREDIT_AMOUNT + Rc.DEBIT_AMOUNT)

Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   GridEX1.Visible = False
   
   Call InitGrid1
   GridEX1.Visible = True

   Call GetTotalPriceReceipt
   GridEX1.ItemCount = CountItem(m_SumBill.RcpCnDnItems)
   GridEX1.Rebind
   
End Sub
Private Sub txtSumBillDesc_Change()
   m_HasModify = True
End Sub
Private Sub GetTotalPriceReceipt()
Dim II As CBillDetail

Dim Sum3 As Double
Dim Sum4 As Double
Dim Sum5 As Double

   Sum3 = 0
   Sum4 = 0
   Sum5 = 0
   
   For Each II In m_SumBill.RcpCnDnItems
      If II.Flag <> "D" Then
            Sum3 = Sum3 + II.PAID_AMOUNT
            Sum4 = Sum4 + II.DEBIT_AMOUNT
            Sum5 = Sum5 + II.CREDIT_AMOUNT
      End If
   Next II
   Set II = Nothing
   
   
   txtPaidAmount.Text = Format(Sum3, "0.00")
   txtDebitAmount.Text = Format(Sum4, "0.00")
   txtCreditAmount.Text = Format(Sum5, "0.00")
   txtAfterDebitCredit.Text = Format(Sum3 + Sum4 - Sum5, "0.00")
   
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - (2 * GridEX1.Left)
   SSFrame4.Width = ScaleWidth
   SSFrame4.Top = ScaleHeight - 1350
   TabStrip1.Width = GridEX1.Width
   GridEX1.Height = SSFrame4.Top - GridEX1.Top - 40
   cmdAdd.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdPrint.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
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

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub
