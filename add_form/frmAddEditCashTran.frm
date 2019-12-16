VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCashTran 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAddEditCashTran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   11880
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBAccountType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7220
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3150
         Width           =   1815
      End
      Begin VB.ComboBox cboPaymentType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   900
         Width           =   3495
      End
      Begin PorkShop.uctlTextLookup uctlChequeType 
         Height          =   405
         Left            =   1860
         TabIndex        =   4
         Top             =   2700
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlDate uctlChequeDate 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtChequeNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1350
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtChequeAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   8
         Top             =   4560
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlEffectiveDate 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   2250
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlBank 
         Height          =   405
         Left            =   1860
         TabIndex        =   5
         Top             =   3630
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextLookup uctlBankBranch 
         Height          =   405
         Left            =   1860
         TabIndex        =   6
         Top             =   4080
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextLookup uctlBankAccountLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   7
         Top             =   3150
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtFeeAmount 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   5040
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtRealRcp 
         Height          =   435
         Left            =   6300
         TabIndex        =   28
         Top             =   5040
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin VB.Label Label4 
         Caption         =   "Label1"
         Height          =   435
         Left            =   8520
         TabIndex        =   30
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label lblRealRcp 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   29
         Top             =   5100
         Width           =   1575
      End
      Begin VB.Label lblFeeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   26
         Top             =   5100
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4080
         TabIndex        =   25
         Top             =   5100
         Width           =   1575
      End
      Begin VB.Label lblBankAccount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   24
         Top             =   3330
         Width           =   1575
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2205
         TabIndex        =   10
         Top             =   5760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashTran.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblPaymentType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   23
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4050
         TabIndex        =   22
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   21
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   20
         Top             =   3750
         Width           =   1575
      End
      Begin VB.Label lblChequeType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   19
         Top             =   2820
         Width           =   1575
      End
      Begin VB.Label lblEffectiveDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label lblChequeDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label lblChequeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label lblChequeNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   1410
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5505
         TabIndex        =   12
         Top             =   5760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3855
         TabIndex        =   11
         Top             =   5760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashTran.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCashTran"
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
Public Area As Long
Public AparMasID As Long
Public DocumentType As SELL_BILLING_DOCTYPE
Public ItemAmount As String

Private m_Cheque As CCheque
Private Mr As CMasterRef
Private m_ApArMas As CAPARMas

Public TempCollection As Collection
Private m_ChequeTypes As Collection
Private m_ApAr As Collection
Private m_Banks As Collection
Private m_BankBranchs As Collection
Private m_BankAccounts As Collection
Private Sub cboBAccountType_Click()
   m_HasModify = True
End Sub
Private Sub cboPaymentType_Click()
Dim TempID As Long
   
   m_HasModify = True
   
   TempID = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   If TempID = CASH_PMT Then
      txtChequeNo.Enabled = False
      txtChequeNo.Text = ""
      uctlChequeDate.Enable = False
      uctlEffectiveDate.Enable = False
      uctlChequeDate.ShowDate = -1
      uctlEffectiveDate.ShowDate = -1
      uctlChequeType.Enabled = False
      uctlBank.Enabled = False
      uctlBank.MyTextBox.Text = ""
      uctlBank.MyCombo.ListIndex = -1
      uctlBankBranch.Enabled = False
      uctlBankBranch.MyTextBox.Text = ""
      uctlBankBranch.MyCombo.ListIndex = -1
      cboBAccountType.Enabled = False
      uctlBankAccountLookup.Enabled = False
      uctlBankAccountLookup.MyTextBox.Text = ""
      uctlBankAccountLookup.MyCombo.ListIndex = -1
      txtChequeAmount.Enabled = True
      txtFeeAmount.Enabled = False
      txtFeeAmount.Text = ""
      'uctlChequeDate.TabStop = False
      'uctlEffectiveDate.TabStop = False
   ElseIf TempID = BANKTRF_PMT Then
      txtChequeNo.Enabled = False
      txtChequeNo.Text = ""
      uctlChequeDate.Enable = False
      uctlEffectiveDate.Enable = False
      uctlChequeDate.ShowDate = -1
      uctlEffectiveDate.ShowDate = -1
      uctlChequeType.Enabled = False
      uctlBank.Enabled = False
      uctlBankBranch.Enabled = False
      cboBAccountType.Enabled = False
      uctlBankAccountLookup.Enabled = True
      txtChequeAmount.Enabled = True
      txtFeeAmount.Enabled = True
      'uctlChequeDate.TabStop = False
      'uctlEffectiveDate.TabStop = False
   ElseIf TempID = CHEQUE_HAND_PMT Or TempID = CHEQUE_BANK_PMT Then
      txtChequeNo.Enabled = True
      uctlChequeDate.Enable = True
      uctlEffectiveDate.Enable = True
      uctlChequeDate.ShowDate = Now
      uctlEffectiveDate.ShowDate = Now
      uctlChequeType.Enabled = True
      uctlBank.Enabled = False
      uctlBankBranch.Enabled = False
      cboBAccountType.Enabled = False
      If Area = 1 And TempID = CHEQUE_HAND_PMT Then
         uctlBankAccountLookup.Enabled = False 'เช็ครับไม่ต้องระบุสมุดบัญชีแต่จะทำใบ pay in ทีหลัง
      ElseIf Area = 2 Or TempID = CHEQUE_BANK_PMT Then
         uctlBankAccountLookup.Enabled = True 'เช็คจ่ายต้องระบุสมุดบัญชีว่าตัดจากบัญชีใด   หรือ จาก เช็คเข้าธนาคารโดยตรง
      End If
      txtChequeAmount.Enabled = True
      If TempID = CHEQUE_BANK_PMT Then
         txtFeeAmount.Enabled = True
      Else
         txtFeeAmount.Enabled = False
      End If
   End If
End Sub

Private Sub cboPaymentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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
         
         Call ParentForm.RefreshCashTran
         Exit Sub
      End If
      
      ID = NewID
      Call QueryData(True)
   ElseIf ShowMode = SHOW_ADD Then
      cboPaymentType.ListIndex = -1
      txtChequeNo.Text = ""
      uctlChequeDate.ShowDate = -1
      uctlEffectiveDate.ShowDate = -1
      uctlChequeType.MyCombo.ListIndex = -1
      uctlBank.MyCombo.ListIndex = -1
      uctlBankBranch.MyCombo.ListIndex = -1
      uctlBankAccountLookup.MyCombo.ListIndex = -1
      txtChequeAmount.Text = ""
      txtFeeAmount.Text = ""
   End If
   
   cboPaymentType.SetFocus
   Call ParentForm.RefreshCashTran
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
      
      Dim Tm As CCashTransferItem
      Set Tm = TempCollection.Item(ID)
      
      PaymentType = Tm.ImportItemEx.GetFieldValue("PAYMENT_TYPE")
      cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, PaymentType)
      txtChequeNo.Text = Tm.ImportItemEx.Cheque.GetFieldValue("CHEQUE_NO")
      txtChequeAmount.Text = Tm.ImportItemEx.GetFieldValue("AMOUNT")
      If PaymentType = CHEQUE_BANK_PMT Then
         txtFeeAmount.Text = Tm.ImportItem.GetFieldValue("FEE_AMOUNT")
      Else
         txtFeeAmount.Text = Tm.ImportItemEx.GetFieldValue("FEE_AMOUNT")
      End If
      uctlChequeDate.ShowDate = Tm.ImportItemEx.Cheque.GetFieldValue("CHEQUE_DATE")
      uctlEffectiveDate.ShowDate = Tm.ImportItemEx.Cheque.GetFieldValue("EFFECTIVE_DATE")
      uctlChequeType.MyCombo.ListIndex = IDToListIndex(uctlChequeType.MyCombo, Tm.ImportItemEx.Cheque.GetFieldValue("CHEQUE_TYPE"))
      If PaymentType = BANKTRF_PMT Then
         uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, Tm.ImportItemEx.GetFieldValue("BANK_ID"))
         uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, Tm.ImportItemEx.GetFieldValue("BANK_BRANCH"))
         uctlBankAccountLookup.MyCombo.ListIndex = IDToListIndex(uctlBankAccountLookup.MyCombo, Tm.ImportItemEx.GetFieldValue("BANK_ACCOUNT"))
      ElseIf PaymentType = CHEQUE_HAND_PMT Then
         uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, Tm.ImportItemEx.Cheque.GetFieldValue("BANK_ID"))
         uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, Tm.ImportItemEx.Cheque.GetFieldValue("BANK_BRANCH"))
      ElseIf PaymentType = CASH_PMT Then
         uctlBank.MyCombo.ListIndex = -1
         uctlBankBranch.MyCombo.ListIndex = -1
         uctlBankAccountLookup.MyCombo.ListIndex = -1
      ElseIf PaymentType = CHEQUE_BANK_PMT Then
         uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, Tm.ImportItem.GetFieldValue("BANK_ID"))
         uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, Tm.ImportItem.GetFieldValue("BANK_BRANCH"))
         uctlBankAccountLookup.MyCombo.ListIndex = IDToListIndex(uctlBankAccountLookup.MyCombo, Tm.ImportItem.GetFieldValue("BANK_ACCOUNT"))
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim PaymentType As PAYMENT_TYPE
   
   PaymentType = cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))
   
   If Not VerifyCombo(lblPaymentType, cboPaymentType, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblChequeNo, txtChequeNo, Not txtChequeNo.Enabled) Then
      Exit Function
   End If
   If Not VerifyDate(lblChequeDate, uctlChequeDate, Not txtChequeNo.Enabled) Then
      Exit Function
   End If
   If Not VerifyDate(lblEffectiveDate, uctlEffectiveDate, Not txtChequeNo.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBank, uctlBank.MyCombo, Not uctlBank.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankBranch, uctlBankBranch.MyCombo, Not uctlBankBranch.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankAccount, uctlBankAccountLookup.MyCombo, Not uctlBankAccountLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblChequeAmount, txtChequeAmount, Not txtChequeAmount.Enabled) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(USERNAME_UNIQUE, txtChequeNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtChequeNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   
   Dim CashTranItem As CCashTransferItem
   Dim IIEx As CCashTran
   Dim Ei As CCashTran
   Dim II As CCashTran
   
   If ShowMode = SHOW_ADD Then
      Set IIEx = New CCashTran
      Set CashTranItem = New CCashTransferItem
      IIEx.Flag = "A"
      CashTranItem.Flag = "A"
      Set CashTranItem.ImportItemEx = IIEx
      If PaymentType = CHEQUE_BANK_PMT Then
         Set Ei = New CCashTran
         Set II = New CCashTran
         Ei.Flag = "A"
         II.Flag = "A"
         Set CashTranItem.ExportItem = Ei
         Set CashTranItem.ImportItem = II
      End If
      Call TempCollection.add(CashTranItem)
   Else     'If Not(ShowMode = SHOW_ADD) Then
      Set CashTranItem = TempCollection.Item(ID)
      If CashTranItem.Flag <> "A" Then
         CashTranItem.Flag = "E"
         CashTranItem.ImportItemEx.Flag = "E"
         Set IIEx = CashTranItem.ImportItemEx
         If PaymentType = CHEQUE_BANK_PMT Then
            If IIEx.GetFieldValue("OLD_PAYMENT_TYPE") <> CHEQUE_BANK_PMT Then
               Set Ei = New CCashTran
               Set II = New CCashTran
               Ei.Flag = "A"
               II.Flag = "A"
               Set CashTranItem.ExportItem = Ei
               Set CashTranItem.ImportItem = II
            Else
               CashTranItem.ExportItem.Flag = "E"
               CashTranItem.ImportItem.Flag = "E"
               Set Ei = CashTranItem.ExportItem
               Set II = CashTranItem.ImportItem
            End If
         Else           'If Not(PaymentType = CHEQUE_BANK_PMT) Then
            If IIEx.GetFieldValue("OLD_PAYMENT_TYPE") = CHEQUE_BANK_PMT Then
               Set Ei = CashTranItem.ExportItem
               Set II = CashTranItem.ImportItem
               If Not (Ei Is Nothing) Then
                  Ei.Flag = "D"
               End If
               If Not (II Is Nothing) Then
                  II.Flag = "D"
               End If
            End If
         End If
      Else     'CashTranItem.Flag = "A"
         Set IIEx = CashTranItem.ImportItemEx
         If PaymentType = CHEQUE_BANK_PMT Then
            Set Ei = CashTranItem.ExportItem
            Set II = CashTranItem.ImportItem
            If Ei Is Nothing Then
               Set Ei = New CCashTran
               Ei.Flag = "A"
               Set CashTranItem.ExportItem = Ei
            End If
            If II Is Nothing Then
               Set II = New CCashTran
               II.Flag = "A"
               Set CashTranItem.ImportItem = II
            End If
         Else
            Set Ei = CashTranItem.ExportItem
            Set II = CashTranItem.ImportItem
            If Not (Ei Is Nothing) Then
               Ei.Flag = "D"
            End If
            If Not (II Is Nothing) Then
               II.Flag = "D"
            End If
         End If
      End If
   End If
   
   Call IIEx.SetFieldValue("PAYMENT_TYPE", PaymentType)
   Call IIEx.SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(cboPaymentType.ItemData(Minus2Zero(cboPaymentType.ListIndex))))
   Call IIEx.SetFieldValue("AMOUNT", Val(txtChequeAmount.Text))
   Call IIEx.SetFieldValue("FROM_BILLING", "Y")
   If PaymentType = CHEQUE_BANK_PMT Then
      Call IIEx.SetFieldValue("FEE_AMOUNT", 0)
      Call IIEx.SetFieldValue("NET_AMOUNT", Val(txtChequeAmount.Text))
   Else
      Call IIEx.SetFieldValue("FEE_AMOUNT", Val(txtFeeAmount.Text))
      Call IIEx.SetFieldValue("NET_AMOUNT", Val(txtChequeAmount.Text) - Val(txtFeeAmount.Text))
   End If
   Call IIEx.SetFieldValue("APAR_MAS_ID", AparMasID)
   
   If PaymentType = CASH_PMT Then
      Call IIEx.SetFieldValue("STEP_ID", 1)
      Call IIEx.SetFieldValue("BANK_ID", -1)
      Call IIEx.SetFieldValue("BANK_BRANCH", -1)
      Call IIEx.SetFieldValue("BANK_ACCOUNT", -1)
   ElseIf PaymentType = BANKTRF_PMT Then
      Call IIEx.SetFieldValue("STEP_ID", 1)
      Call IIEx.SetFieldValue("BANK_ID", uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex)))
      Call IIEx.SetFieldValue("BANK_BRANCH", uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex)))
      Call IIEx.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
      Call IIEx.SetFieldValue("BANK_NAME", uctlBank.MyCombo.Text)
      Call IIEx.SetFieldValue("BRANCH_NAME", uctlBankBranch.MyCombo.Text)
      Call IIEx.SetFieldValue("ACCOUNT_NAME", uctlBankAccountLookup.MyCombo.Text)
      
   ElseIf PaymentType = CHEQUE_HAND_PMT Then
      Call IIEx.SetFieldValue("STEP_ID", 1)
      Call IIEx.Cheque.SetFieldValue("BANK_ID", -1)
      Call IIEx.Cheque.SetFieldValue("BANK_BRANCH", -1)
      Call IIEx.SetFieldValue("BANK_ACCOUNT", -1)
   ElseIf PaymentType = CHEQUE_BANK_PMT Then
      Call IIEx.SetFieldValue("STEP_ID", 3)
      Call IIEx.Cheque.SetFieldValue("BANK_ID", -1)
      Call IIEx.Cheque.SetFieldValue("BANK_BRANCH", -1)
      Call IIEx.SetFieldValue("BANK_ACCOUNT", -1)
      
      Call IIEx.SetFieldValue("BANK_ID", -1)
      Call IIEx.SetFieldValue("BANK_BRANCH", -1)
   End If
   
   Call IIEx.SetFieldValue("TX_TYPE", "I")
   Call IIEx.Cheque.SetFieldValue("DIRECTION", 1)
   
   Call IIEx.Cheque.SetFieldValue("CHEQUE_NO", txtChequeNo.Text)
   Call IIEx.Cheque.SetFieldValue("CHEQUE_AMOUNT", Val(txtChequeAmount.Text))
   Call IIEx.Cheque.SetFieldValue("CHEQUE_DATE", uctlChequeDate.ShowDate)
   Call IIEx.Cheque.SetFieldValue("EFFECTIVE_DATE", uctlEffectiveDate.ShowDate)
   Call IIEx.Cheque.SetFieldValue("CHEQUE_TYPE", uctlChequeType.MyCombo.ItemData(Minus2Zero(uctlChequeType.MyCombo.ListIndex)))
   Call IIEx.Cheque.SetFieldValue("APAR_MAS_ID", AparMasID)
   Call IIEx.Cheque.SetFieldValue("CHEQUE_STATUS", 1)
   If PaymentType = CHEQUE_BANK_PMT Then
      Call IIEx.Cheque.SetFieldValue("BANK_FLAG", "Y")
      Call IIEx.Cheque.SetFieldValue("POST_FLAG", "Y")
   End If
   
   If PaymentType = CHEQUE_BANK_PMT Then
      'ประเภทการจ่ายเงินแบบ เช็คเข้าธนาคารทันทีประกอบด้วย 1.เช็คเข้ามือ 2.เช็คออกจากมือ 3.เงินสดเข้าธนาคาร
      Call Ei.SetFieldValue("STEP_ID", 3)
      Call Ei.SetFieldValue("PAYMENT_TYPE", CHEQUE_BANK_PMT)
      Call Ei.SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(CHEQUE_BANK_PMT))
      Call Ei.SetFieldValue("AMOUNT", Val(txtChequeAmount.Text))
      Call Ei.SetFieldValue("FEE_AMOUNT", 0)
      Call Ei.SetFieldValue("NET_AMOUNT", Val(txtChequeAmount.Text))
      Call Ei.SetFieldValue("APAR_MAS_ID", AparMasID)
      Call Ei.SetFieldValue("TX_TYPE", "E")
      
      Call Ei.Cheque.SetFieldValue("CHEQUE_NO", txtChequeNo.Text)
      Call Ei.Cheque.SetFieldValue("CHEQUE_AMOUNT", Val(txtChequeAmount.Text))
      Call Ei.Cheque.SetFieldValue("CHEQUE_DATE", uctlChequeDate.ShowDate)
      Call Ei.Cheque.SetFieldValue("EFFECTIVE_DATE", uctlEffectiveDate.ShowDate)
      Call Ei.Cheque.SetFieldValue("CHEQUE_TYPE", uctlChequeType.MyCombo.ItemData(Minus2Zero(uctlChequeType.MyCombo.ListIndex)))
      Call Ei.Cheque.SetFieldValue("APAR_MAS_ID", AparMasID)
      Call Ei.Cheque.SetFieldValue("CHEQUE_STATUS", 1)
      
      Call II.SetFieldValue("STEP_ID", 3)
      Call II.SetFieldValue("PAYMENT_TYPE", CASH_PMT)
      Call II.SetFieldValue("PAYMENT_TYPE_NAME", PaymentTypeToText(CASH_PMT))
      Call II.SetFieldValue("AMOUNT", Val(txtChequeAmount.Text))
      Call II.SetFieldValue("FEE_AMOUNT", Val(txtFeeAmount.Text))
      Call II.SetFieldValue("NET_AMOUNT", Val(txtChequeAmount.Text) - Val(txtFeeAmount.Text))
      Call II.SetFieldValue("APAR_MAS_ID", AparMasID)
      Call II.SetFieldValue("TX_TYPE", "I")
      Call II.SetFieldValue("BANK_ID", uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex)))
      Call II.SetFieldValue("BANK_BRANCH", uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex)))
      Call II.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
      Call II.SetFieldValue("BANK_NAME", uctlBank.MyCombo.Text)
      Call II.SetFieldValue("BRANCH_NAME", uctlBankBranch.MyCombo.Text)
      Call II.SetFieldValue("ACCOUNT_NAME", uctlBankAccountLookup.MyCombo.Text)
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlChequeType.MyCombo, m_ChequeTypes, , , MASTER_CHEQUE_TYPE)
      Set uctlChequeType.MyCollection = m_ChequeTypes
      
      Call LoadMaster(uctlBankAccountLookup.MyCombo, m_BankAccounts, , , MASTER_BANK_ACCOUNT)
      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
      
      Call LoadMaster(cboBAccountType, , , , MASTER_BACCOUNT_TYPE)
      
      Call LoadMaster(uctlBank.MyCombo, m_Banks, , , MASTER_BANK)
      Set uctlBank.MyCollection = m_Banks
      
      Call LoadMaster(uctlBankBranch.MyCombo, m_BankBranchs, , , MASTER_BBRANCH)
      Set uctlBankBranch.MyCollection = m_BankBranchs
      
      Call InitPaymentType(cboPaymentType)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
         m_HasModify = False
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         uctlChequeDate.ShowDate = Now
         uctlEffectiveDate.ShowDate = Now
         If DocumentType = RECEIPT1_DOCTYPE Then
            cboPaymentType.ListIndex = IDToListIndex(cboPaymentType, CASH_PMT)
            txtChequeAmount.SetFocus
         End If
         txtChequeAmount.Text = ItemAmount
      End If
      
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

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblChequeNo, MapText("เลขที่เช็ค"))
   Call InitNormalLabel(lblChequeAmount, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblChequeDate, MapText("วันที่เช็ค"))
   Call InitNormalLabel(lblEffectiveDate, MapText("วันที่ขึ้นเงิน"))
   Call InitNormalLabel(lblChequeType, MapText("ประเภทเช็ค"))
   Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
   Call InitNormalLabel(lblBankBranch, MapText("สาขาธนาคาร"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblPaymentType, MapText("การชำระเงิน"))
   Call InitNormalLabel(lblBankAccount, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblFeeAmount, MapText("ค่าธรรมเนียม"))
   Call InitNormalLabel(lblRealRcp, MapText("รับจริง"))
   
   Call txtChequeNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtChequeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtFeeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtRealRcp.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Call InitCombo(cboPaymentType)
   Call InitCombo(cboBAccountType)
   
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
   
   Set m_Cheque = New CCheque
   Set Mr = New CMasterRef
   Set m_ApArMas = New CAPARMas
   
   Set m_ChequeTypes = New Collection
   Set m_ApAr = New Collection
   Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
   Set m_BankAccounts = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Cheque = Nothing
   Set Mr = Nothing
   
   Set m_ChequeTypes = Nothing
   Set m_ApAr = Nothing
   Set m_Banks = Nothing
   Set m_BankBranchs = Nothing
   Set m_ApArMas = Nothing
   Set m_BankAccounts = Nothing
   
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
   Set m_Rs = Nothing
   
End Sub
Private Sub txtChequeAmount_Change()
   txtRealRcp.Text = Val(txtChequeAmount.Text) - Val(txtFeeAmount.Text)
   m_HasModify = True
End Sub

Private Sub txtChequeNo_Change()
   m_HasModify = True
End Sub
Private Sub txtFeeAmount_Change()
   txtRealRcp.Text = Val(txtChequeAmount.Text) - Val(txtFeeAmount.Text)
   m_HasModify = True
End Sub
Private Sub uctlBank_Change()
'Dim TempID As Long
'
'   TempID = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
'
'   If TempID > 0 Then
'      Set Mr = Nothing
'      Set Mr = New CMasterRef
'      Call Mr.SetFieldValue("KEY_ID", -1)
'      Call Mr.SetFieldValue("MASTER_AREA", MASTER_BBRANCH)
'      Call Mr.SetFieldValue("PARENT_ID", TempID)
'      Call LoadMaster(Mr, uctlBankBranch.MyCombo, m_BankBranchs)
'      Set uctlBankBranch.MyCollection = m_BankBranchs
'   End If
   
   m_HasModify = True
End Sub

Private Sub uctlBankAccountLookup_Change()
Dim TempID1 As Long
   
   TempID1 = uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex))
   If TempID1 > 0 Then
      Set Mr = GetObject("CMasterRef", m_BankAccounts, Trim(Str(TempID1)))
      cboBAccountType.ListIndex = IDToListIndex(cboBAccountType, Mr.PARENT_ID)
      uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, Mr.PARENT_EX_ID4)
      uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, Mr.PARENT_EX_ID5)
   Else
      cboBAccountType.ListIndex = -1
      uctlBank.MyCombo.ListIndex = -1
      uctlBankBranch.MyCombo.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlBankBranch_Change()
'Dim TempID1 As Long
'Dim TempID2 As Long
'
'   TempID1 = uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex))
'   TempID2 = uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex))
'
'   If TempID2 > 0 Then
'      Set Mr = Nothing
'      Set Mr = New CMasterRef
'      Call Mr.SetFieldValue("KEY_ID", -1)
'      Call Mr.SetFieldValue("MASTER_AREA", MASTER_BANK_ACCOUNT)
'      Call Mr.SetFieldValue("PARENT_EX_ID4", TempID1)
'      Call Mr.SetFieldValue("PARENT_EX_ID5", TempID2)
'      Call LoadMaster(Mr, uctlBankAccountLookup.MyCombo, m_BankAccounts)
'      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
'   End If
   
   m_HasModify = True
End Sub

Private Sub uctlChequeDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlChequeType_Change()
   m_HasModify = True
End Sub

Private Sub uctlEffectiveDate_HasChange()
   m_HasModify = True
End Sub
