VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCopyApArBranch 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmCopyApArBranch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6429
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboCustomerAddress 
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
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   8685
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1980
         TabIndex        =   4
         Top             =   2520
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1980
         TabIndex        =   5
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin PorkShop.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1980
         TabIndex        =   0
         Top             =   840
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlToCustomerLookup 
         Height          =   435
         Left            =   1980
         TabIndex        =   1
         Top             =   1320
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomerAddress 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -480
         TabIndex        =   14
         Top             =   1950
         Width           =   2325
      End
      Begin VB.Label lblToCustomer 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   1005
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   2
         Top             =   2940
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCopyApArBranch.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   10
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   8
         Top             =   3000
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   3
         Top             =   2940
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCopyApArBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public ProcessMode As Long

Private m_Apm  As CAPARMas
Private m_Adr As CAddress
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
Dim PartItemID As Long
   If Not VerifyCombo(lblCustomer, uctlCustomerLookUp.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblToCustomer, uctlToCustomerLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyCombo(lblCustomerAddress, cboCustomerAddress, False) Then
      Exit Sub
   End If

   Call glbDaily.StartTransaction
   
   Me.Enabled = False
   
   Status = CopyBranch(uctlCustomerLookUp.MyCombo.ItemData(Minus2Zero(uctlCustomerLookUp.MyCombo.ListIndex)), uctlToCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlToCustomerLookup.MyCombo.ListIndex)))
   
   Me.Enabled = True
   
   If Status Then
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
      glbErrorLog.ShowUserError
   Else
      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
      glbErrorLog.ShowUserError
   End If
   
   Call cmdOK_Click
   Exit Sub
   
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      m_Apm.APAR_IND = 1
      Call LoadApArMas(m_Apm, uctlCustomerLookUp.MyCombo, m_CustomerColl)
      Set uctlCustomerLookUp.MyCollection = m_CustomerColl
      uctlCustomerLookUp.Visible = True
      
      Call LoadApArMas(m_Apm, uctlToCustomerLookup.MyCombo, m_CustomerColl)
      Set uctlToCustomerLookup.MyCollection = m_CustomerColl
      uctlToCustomerLookup.Visible = True
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
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

Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblCustomer, "จากลูกค้า")
   Call InitNormalLabel(lblToCustomer, "ไปยังลูกค้า")
   Call InitNormalLabel(lblCustomerAddress, MapText("ที่อยู่สาขา"))
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   Call InitCombo(cboCustomerAddress)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
  ' cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
  ' Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
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
      
   Set m_Apm = New CAPARMas
   Set m_Adr = New CAddress
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Apm = Nothing
   Set m_Adr = Nothing
End Sub
Public Function CopyBranch(FromCustomer As Long, ToCustomer As Long) As Boolean
On Error GoTo ErrorHandler
Dim m_Rs As ADODB.Recordset
Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim PERCENT As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean
   
   If Not (prgProgress Is Nothing) Then
      prgProgress.Max = 100
      prgProgress.Min = 0
   End If
   
   HasBegin = True
   
   Set m_Rs = New ADODB.Recordset
   
   Dim Br As CMasterRef
   Set Br = New CMasterRef
   Br.PARENT_EX_ID2 = FromCustomer
   Call Br.QueryData(1, m_Rs, iCount)
   Set Br = Nothing
   
   I = 0
   While Not m_Rs.EOF
      I = I + 1
      prgProgress.Value = MyDiff(I, m_Rs.RecordCount) * 100
      txtPercent.Text = FormatNumber(prgProgress.Value)
      prgProgress.Refresh
      txtPercent.Refresh
      DoEvents
      
      Set Br = New CMasterRef
      Call Br.PopulateFromRS(1, m_Rs)
      
      Br.KEY_CODE = Br.KEY_CODE & "***"
      Br.KEY_NAME = Br.KEY_NAME & "***"
      Br.PARENT_EX_ID2 = ToCustomer
      Br.PARENT_EX_ID3 = cboCustomerAddress.ItemData(Minus2Zero(cboCustomerAddress.ListIndex))
      Br.ShowMode = SHOW_ADD
      Call Br.AddEditData
         
      Set Br = Nothing

      m_Rs.MoveNext
   Wend
   Set Br = Nothing
      
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   If Not (prgProgress Is Nothing) Then
      prgProgress.Value = 100
      txtPercent.Text = 100
   End If
   
   Set m_Rs = Nothing
   
   CopyBranch = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
   End If
   
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.RoutineName = "CopyBranch"
   glbErrorLog.ModuleName = "FrmProcessCommit"
   glbErrorLog.LocalErrorMsg = "Eror"
   glbErrorLog.ShowErrorLog (LOG_MSGBOX)
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   CopyBranch = False
End Function

Private Sub uctlToCustomerLookup_Change()
      If (uctlToCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlToCustomerLookup.MyCombo.ListIndex)) > 0) Then
         Call m_Adr.SetFieldValue("APAR_MAS_ID", uctlToCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlToCustomerLookup.MyCombo.ListIndex)))
         Call LoadAparMasAddress(m_Adr, cboCustomerAddress, , True)
      End If
End Sub
