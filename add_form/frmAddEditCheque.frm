VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCheque 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAddEditCheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10186
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextLookup uctlChequeType 
         Height          =   405
         Left            =   1860
         TabIndex        =   3
         Top             =   2370
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlDate uctlChequeDate 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtChequeNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   11
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
         TabIndex        =   7
         Top             =   4170
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlEffectiveDate 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlBank 
         Height          =   405
         Left            =   1860
         TabIndex        =   4
         Top             =   2820
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextLookup uctlBankBranch 
         Height          =   405
         Left            =   1860
         TabIndex        =   5
         Top             =   3270
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextLookup uctlAPAR 
         Height          =   405
         Left            =   1860
         TabIndex        =   6
         Top             =   3720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   435
         Left            =   4050
         TabIndex        =   20
         Top             =   4230
         Width           =   1575
      End
      Begin VB.Label lblAPAR 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   19
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Top             =   2940
         Width           =   1575
      End
      Begin VB.Label lblChequeType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   2490
         Width           =   1575
      End
      Begin VB.Label lblEffectiveDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   15
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblChequeDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblChequeAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   4230
         Width           =   1575
      End
      Begin VB.Label lblChequeNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4680
         TabIndex        =   9
         Top             =   4890
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3030
         TabIndex        =   8
         Top             =   4890
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCheque.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Cheque As CCheque

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ChequeType As Long

Private Mr As CMasterRef
Private m_ChequeTypes As Collection
Private m_ApAr As Collection
Private m_Banks As Collection
Private m_BankBranchs As Collection
Private m_ApArMas As CAPARMas
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
'   Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
   OKClick = True
   Unload Me
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      Call m_Cheque.SetFieldValue("CHEQUE_ID", ID)
      m_Cheque.QueryFlag = 1
      If Not glbDaily.QueryCheque(m_Cheque, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Cheque.PopulateFromRS(1, m_Rs)
      
      txtChequeNo.Text = m_Cheque.GetFieldValue("CHEQUE_NO")
      txtChequeAmount.Text = m_Cheque.GetFieldValue("CHEQUE_AMOUNT")
      uctlChequeDate.ShowDate = m_Cheque.GetFieldValue("CHEQUE_DATE")
      uctlEffectiveDate.ShowDate = m_Cheque.GetFieldValue("EFFECTIVE_DATE")
      uctlChequeType.MyCombo.ListIndex = IDToListIndex(uctlChequeType.MyCombo, m_Cheque.GetFieldValue("CHEQUE_TYPE"))
      uctlBank.MyCombo.ListIndex = IDToListIndex(uctlBank.MyCombo, m_Cheque.GetFieldValue("BANK_ID"))
      uctlBankBranch.MyCombo.ListIndex = IDToListIndex(uctlBankBranch.MyCombo, m_Cheque.GetFieldValue("BANK_BRANCH"))
      uctlAPAR.MyCombo.ListIndex = IDToListIndex(uctlAPAR.MyCombo, m_Cheque.GetFieldValue("APAR_MAS_ID"))
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
   
   If Not cmdOK.Enabled Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblChequeNo, txtChequeNo, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Cheque.ShowMode = ShowMode
   Call m_Cheque.SetFieldValue("CHEQUE_ID", ID)
   Call m_Cheque.SetFieldValue("CHEQUE_NO", txtChequeNo.Text)
   Call m_Cheque.SetFieldValue("CHEQUE_AMOUNT", Val(txtChequeAmount.Text))
   Call m_Cheque.SetFieldValue("CHEQUE_DATE", uctlChequeDate.ShowDate)
   Call m_Cheque.SetFieldValue("EFFECTIVE_DATE", uctlEffectiveDate.ShowDate)
   Call m_Cheque.SetFieldValue("CHEQUE_TYPE", uctlChequeType.MyCombo.ItemData(Minus2Zero(uctlChequeType.MyCombo.ListIndex)))
   Call m_Cheque.SetFieldValue("BANK_ID", uctlBank.MyCombo.ItemData(Minus2Zero(uctlBank.MyCombo.ListIndex)))
   Call m_Cheque.SetFieldValue("BANK_BRANCH", uctlBankBranch.MyCombo.ItemData(Minus2Zero(uctlBankBranch.MyCombo.ListIndex)))
   Call m_Cheque.SetFieldValue("APAR_MAS_ID", uctlAPAR.MyCombo.ItemData(Minus2Zero(uctlAPAR.MyCombo.ListIndex)))
   Call m_Cheque.SetFieldValue("CHEQUE_STATUS", 1)
   Call m_Cheque.SetFieldValue("DIRECTION", ChequeType)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCheque(m_Cheque, IsOK, True, glbErrorLog) Then
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

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlChequeType.MyCombo, m_ChequeTypes, , , MASTER_CHEQUE_TYPE)
      Set uctlChequeType.MyCollection = m_ChequeTypes
      
      Call LoadMaster(uctlBank.MyCombo, m_Banks, , , MASTER_BANK)
      Set uctlBank.MyCollection = m_Banks
      
      Call LoadMaster(uctlBankBranch.MyCombo, m_BankBranchs, , , MASTER_BBRANCH)
      Set uctlBankBranch.MyCollection = m_BankBranchs
      
      m_ApArMas.APAR_MAS_ID = -1
      m_ApArMas.APAR_IND = ChequeType
      Call LoadApArMas(m_ApArMas, uctlAPAR.MyCombo)
      Set uctlAPAR.MyCollection = m_CustomerColl
      
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
   If ChequeType = 1 Then
      Call InitNormalLabel(lblAPAR, MapText("ลูกค้า"))
   ElseIf ChequeType = 2 Then
      Call InitNormalLabel(lblAPAR, MapText("ผู้ค้า"))
   End If
   
   Call txtChequeNo.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtChequeAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   cmdOK.Enabled = False
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
   
   Set m_Cheque = New CCheque
   Set m_Rs = New ADODB.Recordset
   Set m_Cheque = New CCheque
   Set Mr = New CMasterRef
   
   Set m_ChequeTypes = New Collection
   Set m_ApAr = New Collection
   Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
   Set m_ApArMas = New CAPARMas

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
End Sub
Private Sub txtChequeAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtChequeNo_Change()
   m_HasModify = True
End Sub
Private Sub uctlAPAR_Change()
   m_HasModify = True
End Sub
Private Sub uctlBank_Change()
   m_HasModify = True
End Sub
Private Sub uctlBankBranch_Change()
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
