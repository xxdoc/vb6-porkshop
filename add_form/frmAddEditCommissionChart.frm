VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCommissionChart 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmAddEditCommissionChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3405
      Left            =   0
      TabIndex        =   3
      Top             =   540
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6006
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboParent 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   3855
      End
      Begin Rebate.uctlTextLookup uctlEmployeeLookUp 
         Height          =   435
         Left            =   2520
         TabIndex        =   5
         Top             =   1200
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlCommissionSale 
         Height          =   435
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlCommissionReturn 
         Height          =   435
         Left            =   2520
         TabIndex        =   10
         Top             =   2160
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtOrderID 
         Height          =   435
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   767
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
         TabIndex        =   13
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label lblCommissionReturn 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   2220
         Width           =   2115
      End
      Begin VB.Label lblCommissionSale 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   1740
         Width           =   2115
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   1260
         Width           =   1515
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4065
         TabIndex        =   2
         Top             =   2700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2415
         TabIndex        =   1
         Top             =   2700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCommissionChart.frx":08CA
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
         TabIndex        =   4
         Top             =   810
         Width           =   1605
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
End
Attribute VB_Name = "frmAddEditCommissionChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmAddEditCommissionChart"

Private HasActivate As Boolean
Private m_HasModify As Boolean
Public HeaderText As String
Public OKClick As Boolean
Public ID As Long
Public FK_ID As Long
Public ShowMode As SHOW_MODE_TYPE
Private m_Rs As ADODB.Recordset

Private m_CommissionChart As CCommissionChart
Private m_CommissionCharts As Collection
Private Emp As CEmployee
Private EmpColl As Collection

Private FtSaleColl As Collection
Private FtReturnColl As Collection
Private Sub cboParent_Click()
   m_HasModify = True
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
   
   If Not VerifyCombo(lblEmployee, uctlEmployeeLookUp.MyCombo, False) Then
      Exit Sub
   End If
   
   If Not m_HasModify Then
      Unload Me
      Exit Sub
   End If
   
   If cboParent.ListIndex > 0 Then
      Call m_CommissionChart.SetFieldValue("PARENT_ID", cboParent.ItemData(cboParent.ListIndex))
   Else
      Call m_CommissionChart.SetFieldValue("PARENT_ID", -1)
   End If
   Call m_CommissionChart.SetFieldValue("EMP_ID", uctlEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookUp.MyCombo.ListIndex)))
   Call m_CommissionChart.SetFieldValue("SALE_ID", uctlCommissionSale.MyCombo.ItemData(Minus2Zero(uctlCommissionSale.MyCombo.ListIndex)))
   Call m_CommissionChart.SetFieldValue("RETURN_ID", uctlCommissionReturn.MyCombo.ItemData(Minus2Zero(uctlCommissionReturn.MyCombo.ListIndex)))
   Call m_CommissionChart.SetFieldValue("SALE_NAME", uctlCommissionSale.MyCombo.Text)
   Call m_CommissionChart.SetFieldValue("RETURN_NAME", uctlCommissionReturn.MyCombo.Text)
   Call m_CommissionChart.SetFieldValue("ORDER_ID", Val(txtOrderID.Text))
   
   Call EnableForm(Me, False)
   m_CommissionChart.ShowMode = ShowMode
   Call m_CommissionChart.SetFieldValue("MASTER_FROMTO_ID", FK_ID)
   Call m_CommissionChart.SetFieldValue("COMMISSION_CHART_ID", ID)
   If Not glbDaily.AddEditCommissionChart(m_CommissionChart, IsOK, True, glbErrorLog) Then
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
   glbErrorLog.SystemErrorMsg = err.Description
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
      
      Emp.EMP_ID = -1
      Call LoadEmployee(Emp, uctlEmployeeLookUp.MyCombo)
      Set uctlEmployeeLookUp.MyCollection = m_EmployeeColl
      
      Call LoadCommissionChart(cboParent, m_CommissionCharts, FK_ID)
      
      Call LoadMasterFromTo(uctlCommissionSale.MyCombo, FtSaleColl, COMMISSION_TABLE)
      Set uctlCommissionSale.MyCollection = FtSaleColl
      
      Call LoadMasterFromTo(uctlCommissionReturn.MyCombo, FtReturnColl, RETURN_TABLE)
      Set uctlCommissionReturn.MyCollection = FtReturnColl
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call EnableForm(Me, False)
         Call m_CommissionChart.SetFieldValue("COMMISSION_CHART_ID", ID)
         Call m_CommissionChart.SetFieldValue("MASTER_FROMTO_ID", FK_ID)
         If Not glbDaily.QueryCommissionChart(m_CommissionChart, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         If ItemCount > 0 Then
            Call m_CommissionChart.PopulateFromRS(1, m_Rs)
            
            cboParent.ListIndex = IDToListIndex(cboParent, m_CommissionChart.GetFieldValue("PARENT_ID"))
            uctlEmployeeLookUp.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookUp.MyCombo, m_CommissionChart.GetFieldValue("EMP_ID"))
            uctlCommissionSale.MyCombo.ListIndex = IDToListIndex(uctlCommissionSale.MyCombo, m_CommissionChart.GetFieldValue("SALE_ID"))
            uctlCommissionReturn.MyCombo.ListIndex = IDToListIndex(uctlCommissionReturn.MyCombo, m_CommissionChart.GetFieldValue("RETURN_ID"))
            txtOrderID.Text = m_CommissionChart.GetFieldValue("ORDER_ID")
         End If
         Call EnableForm(Me, True)
         m_HasModify = False
      End If
   End If
   
Call EnableForm(Me, True)
Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      MsgBox Me.Name
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   End If
End Sub
Private Sub Form_Load()
   Set m_Rs = New ADODB.Recordset
   Set m_CommissionChart = New CCommissionChart
   Set Emp = New CEmployee
   Set EmpColl = New Collection
   Set FtSaleColl = New Collection
   Set FtReturnColl = New Collection

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
   Call InitNormalLabel(lblEmployee, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblCommissionSale, MapText("คอมมิตชั่นส่วนขาย"))
   Call InitNormalLabel(lblCommissionReturn, MapText("คอมมิตชั่นส่วนคืน"))
   Call InitNormalLabel(lblOrderID, MapText("ลำดับ"))
   
   Call txtOrderID.SetTextLenType(TEXT_INTEGER, glbSetting.ID_TYPE)
   
   Call InitCombo(cboParent)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdCancel, MapText("ยกเลิก (ESC)"))
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set Emp = Nothing
   Set EmpColl = Nothing
   Set m_CommissionChart = Nothing
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set FtSaleColl = Nothing
   Set FtReturnColl = Nothing
End Sub

Private Sub txtOrderID_Change()
   m_HasModify = True
End Sub

Private Sub uctlCommissionReturn_Change()
   m_HasModify = True
End Sub

Private Sub uctlCommissionSale_Change()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
