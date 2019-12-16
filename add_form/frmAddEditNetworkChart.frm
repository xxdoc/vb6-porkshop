VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditNetworkChart 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmAddEditNetworkChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   10920
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3405
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6006
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboParent 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   8175
      End
      Begin PorkShop.uctlTextLookup uctlCustomerLookUp 
         Height          =   435
         Left            =   2520
         TabIndex        =   2
         Top             =   1560
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtOrderID 
         Height          =   435
         Left            =   2520
         TabIndex        =   0
         Top             =   600
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   767
      End
      Begin Threed.SSCheck ChkHoldFlag 
         Height          =   435
         Left            =   8040
         TabIndex        =   3
         Top             =   1560
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblOrderID 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   840
         TabIndex        =   8
         Top             =   1620
         Width           =   1515
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5625
         TabIndex        =   5
         Top             =   2460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3975
         TabIndex        =   4
         Top             =   2460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditNetworkChart.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblParent 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         TabIndex        =   7
         Top             =   1170
         Width           =   1605
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
End
Attribute VB_Name = "frmAddEditNetworkChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmAddEditNetworkChart"

Private HasActivate As Boolean
Private m_HasModify As Boolean
Public HeaderText As String
Public OKClick As Boolean
Public ID As Long
Public FK_ID As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_Rs As ADODB.Recordset

Public ParentTag As Long

Private m_NetworkChart As CNetworkChart
Private m_NetworkCharts As Collection
Private Customer As CAPARMas
Private CustomerColl As Collection

Private Sub cboParent_Click()
   m_HasModify = True
End Sub

Private Sub cboParent_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub ChkHoldFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub ChkHoldFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorHandler
Dim IsOK As Boolean
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Activate"
   
   If Not VerifyTextControl(lblOrderID, txtOrderID, False) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblParent, cboParent, True) Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblCustomer, uctlCustomerLookUp.MyCombo, False) Then
      Exit Sub
   End If
   If cboParent.ListIndex > 0 Then
      If m_NetworkChart.NETWORK_CHART_ID > 0 And (cboParent.ItemData(cboParent.ListIndex) = m_NetworkChart.NETWORK_CHART_ID) Then
         glbErrorLog.LocalErrorMsg = "ไม่สามารถบันทึกได้เนื่องจาก รายการของลูกข่ายและแม่ข่ายเป็นอันเดียวกัน"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   End If
   
   If Not m_HasModify Then
      Unload Me
      Exit Sub
   End If
   
   If cboParent.ListIndex > 0 Then
      m_NetworkChart.PARENT_ID = cboParent.ItemData(cboParent.ListIndex)
   Else
      m_NetworkChart.PARENT_ID = -1
   End If
   m_NetworkChart.CUSTOMER_ID = uctlCustomerLookUp.MyCombo.ItemData(Minus2Zero(uctlCustomerLookUp.MyCombo.ListIndex))
   m_NetworkChart.ORDER_ID = Val(txtOrderID.Text)
   m_NetworkChart.HOLD_FLAG = Check2Flag(ChkHoldFlag.Value)
   
   
   Call EnableForm(Me, False)
   m_NetworkChart.AddEditMode = ShowMode
   m_NetworkChart.MASTER_VALID_ID = FK_ID
   m_NetworkChart.NETWORK_CHART_ID = ID
   If Not glbDaily.AddEditNetworkChart(m_NetworkChart, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim ItemCount As Long
Dim IsOK As Boolean

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"

   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
      
      Customer.APAR_MAS_ID = -1
      Call LoadApArMas(Customer, uctlCustomerLookUp.MyCombo)
      Set uctlCustomerLookUp.MyCollection = m_CustomerColl
      
      Call LoadNetworkChart(cboParent, m_NetworkCharts, FK_ID)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call EnableForm(Me, False)
         m_NetworkChart.NETWORK_CHART_ID = ID
         m_NetworkChart.MASTER_VALID_ID = FK_ID
         If Not glbDaily.QueryNetworkChart(m_NetworkChart, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If ItemCount > 0 Then
            Call m_NetworkChart.PopulateFromRS(1, m_Rs)
            
            cboParent.ListIndex = IDToListIndex(cboParent, m_NetworkChart.PARENT_ID)
            uctlCustomerLookUp.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookUp.MyCombo, m_NetworkChart.CUSTOMER_ID)
            txtOrderID.Text = m_NetworkChart.ORDER_ID
            ChkHoldFlag.Value = FlagToCheck(m_NetworkChart.HOLD_FLAG)
         End If
         Call EnableForm(Me, True)
         m_HasModify = False
      ElseIf ShowMode = SHOW_ADD Then
         cboParent.ListIndex = IDToListIndex(cboParent, ParentTag)
      End If
   End If
   
Call EnableForm(Me, True)
Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      'Call cmdAddLotItem_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      'Call cmdNext_Click
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
   Set m_Rs = New ADODB.Recordset
   Set m_NetworkChart = New CNetworkChart
   Set Customer = New CAPARMas
   Set CustomerColl = New Collection
   
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
   
   Call InitNormalLabel(lblParent, MapText("ภายใต้"))
   Call InitNormalLabel(lblCustomer, MapText("ลุกค้า"))
   Call InitNormalLabel(lblOrderID, MapText("ลำดับ"))
   
   uctlCustomerLookUp.MyTextBox.SetKeySearch ("CUSTOMER_CODE")
   
   Call txtOrderID.SetTextLenType(TEXT_INTEGER, glbSetting.ID_TYPE)
   
   Call InitCheckBox(ChkHoldFlag, "ระงับการจ่าย")
   Call InitCombo(cboParent)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set Customer = Nothing
   Set CustomerColl = Nothing
   Set m_NetworkChart = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub txtOrderID_Change()
   m_HasModify = True
End Sub
Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub
