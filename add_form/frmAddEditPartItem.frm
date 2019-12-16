VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditPartItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6990
   Icon            =   "frmAddEditPartItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6990
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   8070
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboUnitChange 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   2025
      End
      Begin VB.ComboBox cboUnit 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2400
         Width           =   1755
      End
      Begin VB.ComboBox cboPartType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   4485
      End
      Begin PorkShop.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1470
         Width           =   4485
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtPartNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtUnitAmount 
         Height          =   435
         Left            =   3600
         TabIndex        =   4
         Top             =   2400
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkExpenseFlag 
         Height          =   435
         Left            =   1800
         TabIndex        =   15
         Top             =   3240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2160
         TabIndex        =   7
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPartItem.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3840
         TabIndex        =   8
         Top             =   3840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck ChkExceptionFlag 
         Height          =   435
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblPartType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   13
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   11
         Top             =   1110
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditPartItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PartItem As CStockCode

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Public PartGroupID As Long
Private Sub cboPartType_Click()
   m_HasModify = True
End Sub
Private Sub cboPartType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboUnit_Click()
   m_HasModify = True
   txtUnitAmount.Text = "1"
   cboUnitChange.ListIndex = cboUnit.ListIndex
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboUnitChange_Click()
   m_HasModify = True
End Sub

Private Sub cboUnitChange_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub ChkExceptionFlag_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub ChkExceptionFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkExpenseFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkExpenseFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
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

   If Flag Then
      Call EnableForm(Me, False)
      
      m_PartItem.STOCK_CODE_ID = ID
      m_PartItem.QueryFlag = 1
      If Not glbDaily.QueryStockCode(m_PartItem, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_PartItem.PopulateFromRS(1, m_Rs)
      
      txtName.Text = m_PartItem.STOCK_DESC
      txtPartNo.Text = m_PartItem.STOCK_NO
      cboPartType.ListIndex = IDToListIndex(cboPartType, m_PartItem.STOCK_TYPE)
      cboUnit.ListIndex = IDToListIndex(cboUnit, m_PartItem.UNIT_ID)
      cboUnitChange.ListIndex = IDToListIndex(cboUnitChange, m_PartItem.UNIT_CHANGE_ID)
      txtUnitAmount.Text = m_PartItem.UNIT_AMOUNT
      ChkExceptionFlag.Value = FlagToCheck(m_PartItem.EXCEPTION_FLAG)
      chkExpenseFlag.Value = FlagToCheck(m_PartItem.EXPENSE_FLAG)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblPartNo, txtPartNo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblPartType, cboPartType, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblUnit, cboUnit, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtPartNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_PartItem.ShowMode = ShowMode
   m_PartItem.STOCK_CODE_ID = ID
   m_PartItem.STOCK_NO = txtPartNo.Text
   m_PartItem.STOCK_DESC = txtName.Text
   m_PartItem.STOCK_TYPE = cboPartType.ItemData(Minus2Zero(cboPartType.ListIndex))
   m_PartItem.UNIT_ID = cboUnit.ItemData(Minus2Zero(cboUnit.ListIndex))
   m_PartItem.UNIT_CHANGE_ID = cboUnitChange.ItemData(Minus2Zero(cboUnitChange.ListIndex))
   m_PartItem.UNIT_AMOUNT = Val(txtUnitAmount.Text)
   m_PartItem.EXCEPTION_FLAG = Check2Flag(ChkExceptionFlag.Value)
   m_PartItem.EXPENSE_FLAG = Check2Flag(chkExpenseFlag.Value)
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditStockCode(m_PartItem, IsOK, True, glbErrorLog) Then
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
      
      Call LoadMaster(cboPartType, , , , MASTER_STOCKTYPE, PartGroupID)
      
      Call LoadMaster(cboUnit, , , , MASTER_UNIT)
      
      Call LoadMaster(cboUnitChange, , , , MASTER_UNIT)
            
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
   
   Call InitNormalLabel(lblName, MapText("รายการ"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสคลัง"))
   Call InitNormalLabel(lblPartType, MapText("ประเภทรายการ"))
   Call InitNormalLabel(lblUnit, MapText("หน่วยวัด"))
   
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtPartNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtUnitAmount.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitCombo(cboPartType)
   Call InitCombo(cboUnit)
   Call InitCombo(cboUnitChange)
   
   Call InitCheckBox(ChkExceptionFlag, "ยกเลิก")
   Call InitCheckBox(chkExpenseFlag, "เป็นค่าใช้จ่ายไม่คิด STOCK")
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   Set m_PartItem = New CStockCode
   Set m_Rs = New ADODB.Recordset

   Call EnableForm(Me, False)
   m_HasActivate = False
      
   m_HasActivate = False
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub txtPartNo_Change()
   m_HasModify = True
End Sub
Private Sub txtName_Change()
   m_HasModify = True
End Sub
Private Sub txtUnitAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtPartNo_LostFocus()
   If Not CheckUniqueNs(PARTNO_UNIQUE, txtPartNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtPartNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      txtPartNo.SetFocus
      Exit Sub
   End If
End Sub
