VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmReportConfig 
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmReportConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4980
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8784
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboPaperSize 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cboOrientation 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2880
         Width           =   2355
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtMarginBottom 
         Height          =   435
         Left            =   6000
         TabIndex        =   2
         Top             =   1500
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtMarginTop 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1500
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtMarginLeft 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1950
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtMarginRight 
         Height          =   435
         Left            =   6000
         TabIndex        =   4
         Top             =   1950
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtHeadOffset 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   3300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtDummyOffset 
         Height          =   435
         Left            =   6000
         TabIndex        =   9
         Top             =   3300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtFontSize 
         Height          =   435
         Left            =   6000
         TabIndex        =   6
         Top             =   2460
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtFontName 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2460
         Width           =   1875
         _ExtentX        =   2672
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdFontName 
         Height          =   435
         Left            =   3720
         TabIndex        =   30
         Top             =   2460
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmReportConfig.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblFontSize 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   29
         Top             =   2550
         Width           =   1575
      End
      Begin VB.Label lblHeadOffset 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label lblDummyOffset 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label Label3 
         Height          =   315
         Left            =   3480
         TabIndex        =   26
         Top             =   3330
         Width           =   525
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   7620
         TabIndex        =   25
         Top             =   3360
         Width           =   525
      End
      Begin VB.Label lblCm8 
         Height          =   315
         Left            =   7620
         TabIndex        =   24
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lblCm7 
         Height          =   315
         Left            =   7620
         TabIndex        =   23
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label lblCm3 
         Height          =   315
         Left            =   3480
         TabIndex        =   22
         Top             =   2010
         Width           =   435
      End
      Begin VB.Label lblCm2 
         Height          =   315
         Left            =   3480
         TabIndex        =   21
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lblOrientation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   2970
         Width           =   1575
      End
      Begin VB.Label lblFontName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   2550
         Width           =   1575
      End
      Begin VB.Label lblMarginRight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         TabIndex        =   18
         Top             =   2010
         Width           =   1665
      End
      Begin VB.Label lblMarginLeft 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   2010
         Width           =   1665
      End
      Begin VB.Label lblMarginTop 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label lblMarginBottom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4230
         TabIndex        =   15
         Top             =   1560
         Width           =   1665
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2625
         TabIndex        =   10
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmReportConfig.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4275
         TabIndex        =   12
         Top             =   4080
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblPaperSize 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   1110
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmReportConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_ReportConfig As CReportConfig
Private m_Houses As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ReportKey As String
Public ReportMode As Long

Private FileName As String
Private m_SumUnit As Double
Private m_OldPartItemID As Long
Private m_PigStatus As Collection

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_ReportConfig.SetFieldValue("REPORT_CONFIG_ID", ID)
      If Not glbDaily.QueryReportConfig(m_ReportConfig, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_ReportConfig.PopulateFromRS(1, m_Rs)
      
      txtMarginTop.Text = m_ReportConfig.GetFieldValue("MARGIN_TOP")
      txtMarginBottom.Text = m_ReportConfig.GetFieldValue("MARGIN_BOTTOM")
      txtMarginLeft.Text = m_ReportConfig.GetFieldValue("MARGIN_LEFT")
      txtMarginRight.Text = m_ReportConfig.GetFieldValue("MARGIN_RIGHT")
      cboPaperSize.ListIndex = IDToListIndex(cboPaperSize, m_ReportConfig.GetFieldValue("PAPER_SIZE"))
      cboOrientation.ListIndex = IDToListIndex(cboOrientation, m_ReportConfig.GetFieldValue("Orientation"))
      txtFontName.Text = m_ReportConfig.GetFieldValue("FONT_NAME")
      txtHeadOffset.Text = m_ReportConfig.GetFieldValue("HEAD_OFFSET")
      txtDummyOffset.Text = m_ReportConfig.GetFieldValue("DUMMY_OFFSET")
      txtFontSize.Text = m_ReportConfig.GetFieldValue("FONT_SIZE")
   Else
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If ShowMode = SHOW_ADD Then
   ElseIf ShowMode = SHOW_EDIT Then
   End If

   If Not VerifyCombo(lblPaperSize, cboPaperSize, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginTop, txtMarginTop, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginBottom, txtMarginBottom, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginLeft, txtMarginLeft, (ReportMode <> 1)) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblMarginRight, txtMarginRight, (ReportMode <> 1)) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   
   m_ReportConfig.ShowMode = ShowMode
   Call m_ReportConfig.SetFieldValue("REPORT_KEY", ReportKey)
   Call m_ReportConfig.SetFieldValue("REPORT_CONFIG_ID", ID)
   Call m_ReportConfig.SetFieldValue("MARGIN_TOP", Val(txtMarginTop.Text))
   Call m_ReportConfig.SetFieldValue("MARGIN_BOTTOM", Val(txtMarginBottom.Text))
   Call m_ReportConfig.SetFieldValue("MARGIN_LEFT", Val(txtMarginLeft.Text))
   Call m_ReportConfig.SetFieldValue("MARGIN_RIGHT", Val(txtMarginRight.Text))
   Call m_ReportConfig.SetFieldValue("PAPER_SIZE", cboPaperSize.ItemData(Minus2Zero(cboPaperSize.ListIndex)))
   Call m_ReportConfig.SetFieldValue("Orientation", cboOrientation.ItemData(Minus2Zero(cboOrientation.ListIndex)))
   Call m_ReportConfig.SetFieldValue("FONT_NAME", txtFontName.Text)
   Call m_ReportConfig.SetFieldValue("FONT_SIZE", Val(txtFontSize.Text))
   Call m_ReportConfig.SetFieldValue("HEAD_OFFSET", Val(txtHeadOffset.Text))
   Call m_ReportConfig.SetFieldValue("DUMMY_OFFSET", Val(txtDummyOffset.Text))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditReportConfig(m_ReportConfig, IsOK, True, glbErrorLog) Then
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
Private Sub cboOrientation_Click()
   m_HasModify = True
End Sub

Private Sub cboOrientation_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboPaperSize_Click()
   m_HasModify = True
End Sub

Private Sub cboPaperSize_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdFontName_Click()
   CommonDialog1.Flags = cdlCFPrinterFonts
   CommonDialog1.ShowFont
   
   If Len(CommonDialog1.FontName) > 0 Then
      txtFontName.Text = CommonDialog1.FontName
   End If
   If CommonDialog1.FontSize > 0 Then
      txtFontSize.Text = CommonDialog1.FontSize
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call InitOrientation(cboOrientation)
      Call InitPaperSize(cboPaperSize)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_ReportConfig = Nothing
   Set m_Houses = Nothing
   Set m_Employees = Nothing
   Set m_PigStatus = Nothing
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
      
   Call InitNormalLabel(lblPaperSize, MapText("ขนาดกระดาษ"))
   Call InitNormalLabel(lblMarginTop, MapText("กั้นหน้าบน"))
   Call InitNormalLabel(lblMarginBottom, MapText("กั้นหน้าล่าง"))
   Call InitNormalLabel(lblMarginLeft, MapText("กั้นหน้าซ้าย"))
   Call InitNormalLabel(lblMarginRight, MapText("กั้นหน้าขวา"))
   Call InitNormalLabel(lblFontName, MapText("ชื่อฟอนต์"))
   Call InitNormalLabel(lblOrientation, MapText("การจัดเรียงหน้า"))
   Call InitNormalLabel(lblHeadOffset, MapText("ปรับบน"))
   Call InitNormalLabel(lblDummyOffset, MapText("ปรับซ้าย"))
   Call InitNormalLabel(lblFontSize, MapText("ขนาด"))
   
   Call InitNormalLabel(lblCm2, MapText("ซ.ม."))
   Call InitNormalLabel(lblCm3, MapText("ซ.ม."))
   Call InitNormalLabel(lblCm7, MapText("ซ.ม."))
   Call InitNormalLabel(lblCm8, MapText("ซ.ม."))
   Call InitNormalLabel(Label1, MapText("TWIP"))
   Call InitNormalLabel(Label3, MapText("TWIP"))
   
   Call txtMarginTop.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginTop.Enabled = (ReportMode = 1)
   Call txtMarginBottom.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginBottom.Enabled = (ReportMode = 1)
   Call txtMarginLeft.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginLeft.Enabled = (ReportMode = 1)
   Call txtMarginRight.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtMarginRight.Enabled = (ReportMode = 1)
   Call txtHeadOffset.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtHeadOffset.Enabled = (ReportMode <> 1)
   Call txtDummyOffset.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtDummyOffset.Enabled = (ReportMode <> 1)
   Call txtFontSize.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   txtFontName.Enabled = False
   txtFontSize.Enabled = False
   
   Call InitCombo(cboPaperSize)
   cboPaperSize.Enabled = (ReportMode = 1)
   
   Call InitCombo(cboOrientation)
   cboOrientation.Enabled = (ReportMode = 1)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFontName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdFontName, MapText("F"))
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
   
   If ReportMode <= 0 Then
      ReportMode = 1
   End If
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_ReportConfig = New CReportConfig
   Set m_Houses = New Collection
   Set m_Employees = New Collection
   Set m_PigStatus = New Collection
End Sub
Private Sub txtDummyOffset_Change()
   m_HasModify = True
End Sub

Private Sub txtFontName_Change()
   m_HasModify = True
End Sub

Private Sub txtFontSize_Change()
   m_HasModify = True
End Sub
Private Sub txtHeadOffset_Change()
   m_HasModify = True
End Sub
Private Sub txtMarginBottom_Change()
   m_HasModify = True
End Sub
Private Sub txtMarginFooter_Change()
   m_HasModify = True
End Sub
Private Sub txtMarginHeader_Change()
   m_HasModify = True
End Sub
Private Sub txtMarginLeft_Change()
   m_HasModify = True
End Sub
Private Sub txtMarginRight_Change()
   m_HasModify = True
End Sub
Private Sub txtMarginTop_Change()
   m_HasModify = True
End Sub
Private Sub txtPaperHeight_Change()
   m_HasModify = True
End Sub
Private Sub txtPaperWidth_Change()
   m_HasModify = True
End Sub
Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub
Private Sub uctlTextBox10_Change()
   m_HasModify = True
End Sub
Private Sub uctlTextBox8_Change()
   m_HasModify = True
End Sub
