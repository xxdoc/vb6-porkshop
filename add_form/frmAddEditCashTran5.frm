VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCashTran5 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmAddEditCashTran5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2865
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   5054
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtChequeAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   960
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtFeeAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1410
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblFeeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4050
         TabIndex        =   8
         Top             =   1470
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   645
         TabIndex        =   2
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashTran5.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblChequeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1020
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3945
         TabIndex        =   4
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2295
         TabIndex        =   3
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashTran5.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCashTran5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ParentForm As Object
Public TempCollection As Collection

Public DocumentType As CASH_DOC_TYPE

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
         
         Call ParentForm.RefreshGrid(DocumentType, True)
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      txtChequeAmount.Text = ""
      txtFeeAmount.Text = ""
   End If
   
   Call ParentForm.RefreshGrid(DocumentType, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me

End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim PaymentType As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Dim Ji As CCashTransferItem
      Set Ji = TempCollection.Item(ID)
      
      PaymentType = Ji.ExportItem.GetFieldValue("PAYMENT_TYPE")
      txtChequeAmount.Text = Ji.ImportItem.GetFieldValue("AMOUNT")
      txtFeeAmount.Text = Ji.ImportItem.GetFieldValue("FEE_AMOUNT")
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyTextControl(lblChequeAmount, txtChequeAmount, Not txtChequeAmount.Enabled) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CCashTransferItem
   Dim Ei As CCashTran
   Dim II As CCashTran
   
   If ShowMode = SHOW_ADD Then
      Set Ei = New CCashTran
      Set II = New CCashTran
      Set EnpAddress = New CCashTransferItem

      Ei.Flag = "A"
      II.Flag = "A"
      EnpAddress.Flag = "A"

      Set EnpAddress.ExportItem = Ei
      Set EnpAddress.ImportItem = II

      Call TempCollection.add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
         EnpAddress.ExportItem.Flag = "E"
         EnpAddress.ImportItem.Flag = "E"
      End If
   End If

   'นำฝากเงินสดในมือ
   Call EnpAddress.ExportItem.SetFieldValue("PAYMENT_TYPE", CASH_PMT) 'ออกเป็นเงินสด
   Call EnpAddress.ExportItem.SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(CASH_PMT))
   Call EnpAddress.ExportItem.SetFieldValue("AMOUNT", Val(txtChequeAmount.Text))
   Call EnpAddress.ExportItem.SetFieldValue("TX_TYPE", "E")
   'เงินสดในมือลดลง
   Call EnpAddress.ExportItem.SetFieldValue("BANK_ID", -1)
   Call EnpAddress.ExportItem.SetFieldValue("BANK_BRANCH", -1)
   Call EnpAddress.ExportItem.SetFieldValue("BANK_ACCOUNT", -1)
   Call EnpAddress.ExportItem.SetFieldValue("BANK_NAME", "")
   Call EnpAddress.ExportItem.SetFieldValue("BRANCH_NAME", "")
    
   Call EnpAddress.ImportItem.SetFieldValue("PAYMENT_TYPE", CASH_PMT) 'เข้าเป็นเงินสด
   Call EnpAddress.ImportItem.SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(CASH_PMT))
   Call EnpAddress.ImportItem.SetFieldValue("AMOUNT", Val(txtChequeAmount.Text))
   Call EnpAddress.ImportItem.SetFieldValue("FEE_AMOUNT", Val(txtFeeAmount.Text))
   Call EnpAddress.ImportItem.SetFieldValue("NET_AMOUNT", Val(txtChequeAmount.Text) - Val(txtFeeAmount.Text))
   Call EnpAddress.ImportItem.SetFieldValue("TX_TYPE", "I")

    Call EnpAddress.ImportItem.SetFieldValue("BANK_ID", -1)
    Call EnpAddress.ImportItem.SetFieldValue("BANK_BRANCH", -1)
    Call EnpAddress.ImportItem.SetFieldValue("BANK_ACCOUNT", -1)
      
    'จะเป็นค่าเดียวกันกับ BANK_ID, BANK_BRANCH, BANK_ACCOUNT ของ CASH_DOC
    
   Set EnpAddress = Nothing

   Call EnableForm(Me, True)
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
   ElseIf Shift = 0 And KeyCode = 123 Then
'      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblChequeAmount, MapText("จำนวนเงิน"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblFeeAmount, MapText("ค่าธรรมเนียม"))
   
   Call txtChequeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtFeeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
    
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
     
    
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub txtChequeAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtFeeAmount_Change()
   m_HasModify = True
End Sub
