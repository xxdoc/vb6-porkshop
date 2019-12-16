VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddReceiptEditItemEx 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddReceiptEditItemEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   4895
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextBox txtLeft 
         Height          =   435
         Left            =   9750
         TabIndex        =   0
         Top             =   210
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtCashDiscountPercent 
         Height          =   435
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   975
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtCashDiscountAmount 
         Height          =   435
         Left            =   4680
         TabIndex        =   13
         Top             =   720
         Width           =   1935
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtLeftDiscount 
         Height          =   435
         Left            =   9750
         TabIndex        =   15
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtPaidAmount 
         Height          =   435
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtLeftPaid 
         Height          =   435
         Left            =   4680
         TabIndex        =   19
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   767
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblLeftPaid 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label lblPaidAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lblLeftDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         Top             =   720
         Width           =   2685
      End
      Begin VB.Label lblCashDiscountAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   765
         Width           =   1125
      End
      Begin VB.Label lblCashDiscountPercent 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1485
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   3525
         TabIndex        =   1
         Top             =   1830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptEditItemEx.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5175
         TabIndex        =   2
         Top             =   1830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptEditItemEx.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   6825
         TabIndex        =   3
         Top             =   1830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   8640
         TabIndex        =   6
         Top             =   270
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmAddReceiptEditItemEx"
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

Private m_BillingDoc As CBillingDoc

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form

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
   
   Call InitNormalLabel(lblDocumentNo, MapText("หมายเลข"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่"))
   Call InitNormalLabel(lblLeft, MapText("คงค้าง"))
   Call InitNormalLabel(lblCashDiscountPercent, MapText("ส่วนลดรับ"))
   Call InitNormalLabel(Label1, MapText("%"))
   Call InitNormalLabel(lblCashDiscountAmount, MapText("จำนวน"))
   Call InitNormalLabel(lblLeftDiscount, MapText("จำนวนคงค้างหลังหักส่วนลด"))
   Call InitNormalLabel(lblPaidAmount, MapText("จำนวนที่จ่าย"))
   Call InitNormalLabel(lblLeftPaid, MapText("คงค้าง"))
   
   
   txtDocumentNo.Enabled = False
   uctlDocumentDate.Enable = False
   Call txtLeft.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtLeft.Enabled = False
   Call txtCashDiscountPercent.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtCashDiscountAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtCashDiscountAmount.Enabled = False
   Call txtLeftDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtLeftDiscount.Enabled = False
   Call txtPaidAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtLeftPaid.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtLeftPaid.Enabled = False
   txtCashDiscountPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CRcpCnDn_Item
         
         Set BD = TempCollection.Item(ID)
         
         txtDocumentNo.Text = BD.GetFieldValue("DOC_NO")
         uctlDocumentDate.ShowDate = BD.GetFieldValue("DOC_DATE")
         txtCashDiscountPercent.Text = "0"
         txtCashDiscountAmount.Text = "0"
         txtLeft.Text = BD.GetFieldValue("ITEM_AMOUNT")
         txtPaidAmount.Text = BD.GetFieldValue("PAID_AMOUNT")
         
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
   ElseIf ShowMode = SHOW_ADD Then
     
   End If
   Call QueryData(True)
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

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim BD As CRcpCnDn_Item
   If ShowMode = SHOW_ADD Then
      Set BD = New CRcpCnDn_Item
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" And BD.Flag <> "" Then
         BD.Flag = "E"
      End If
   End If
   
   Call BD.SetFieldValue("PAID_AMOUNT", txtPaidAmount.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      'm_HasModify = False
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
      Call cmdNext_Click
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
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_BillingDoc = New CBillingDoc
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_BillingDoc = Nothing

End Sub

Private Sub txtCashDiscountAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtCashDiscountPercent_Change()
   Call Calculate
   txtPaidAmount.Text = txtLeftDiscount.Text
   Call Calculate
   m_HasModify = True
End Sub

Private Sub txtPaidAmount_Change()
   Call Calculate
   m_HasModify = True
End Sub
Private Sub Calculate()
   txtCashDiscountAmount.Text = FormatNumber(Val(txtLeft.Text) * Val(txtCashDiscountPercent.Text) / 100, , False)
   txtLeftDiscount.Text = FormatNumber(Val(txtLeft.Text) - Val(txtCashDiscountAmount.Text), , False)
   If Val(txtPaidAmount.Text) <= 0 Or Val(txtPaidAmount.Text) >= Val(txtLeftDiscount.Text) Then
      txtPaidAmount.Text = txtLeftDiscount.Text
   End If
   txtLeftPaid.Text = FormatNumber(Val(txtLeftDiscount.Text) - Val(txtPaidAmount.Text), , False)
End Sub
