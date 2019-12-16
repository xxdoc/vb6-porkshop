VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditJournal 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboApAr 
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
         TabIndex        =   3
         Top             =   1500
         Width           =   2265
      End
      Begin Xivess.uctlDate uctlJournalDate 
         Height          =   405
         Left            =   6840
         TabIndex        =   2
         Top             =   1050
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Xivess.uctlTextLookup uctlApArMasLookup 
         Height          =   405
         Left            =   4140
         TabIndex        =   4
         Top             =   1500
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin VB.ComboBox cboDepartment 
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
         TabIndex        =   7
         Top             =   2430
         Width           =   3495
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   10
         Top             =   3210
         Width           =   11595
         _ExtentX        =   20452
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
      Begin Xivess.uctlTextBox txtJournalCode 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtJournalDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   6
         Top             =   1950
         Width           =   9225
         _ExtentX        =   16907
         _ExtentY        =   794
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3975
         Left            =   150
         TabIndex        =   11
         Top             =   3750
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7011
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
         Column(1)       =   "frmAddEditJournal.frx":27A2
         Column(2)       =   "frmAddEditJournal.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditJournal.frx":290E
         FormatStyle(2)  =   "frmAddEditJournal.frx":2A6A
         FormatStyle(3)  =   "frmAddEditJournal.frx":2B1A
         FormatStyle(4)  =   "frmAddEditJournal.frx":2BCE
         FormatStyle(5)  =   "frmAddEditJournal.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditJournal.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtDebit 
         Height          =   435
         Left            =   6540
         TabIndex        =   8
         Top             =   2430
         Width           =   1665
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin Xivess.uctlTextBox txtCredit 
         Height          =   435
         Left            =   9420
         TabIndex        =   9
         Top             =   2430
         Width           =   1665
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8250
         TabIndex        =   25
         Top             =   2550
         Width           =   1065
      End
      Begin VB.Label lblDebit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5370
         TabIndex        =   24
         Top             =   2550
         Width           =   1065
      End
      Begin Threed.SSCheck chkPostFlag 
         Height          =   405
         Left            =   9600
         TabIndex        =   5
         Top             =   1500
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblJournalDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5250
         TabIndex        =   23
         Top             =   1050
         Width           =   1455
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4170
         TabIndex        =   1
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJournal.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblApArMas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   22
         Top             =   1560
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJournal.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   16
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJournal.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJournal.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblJournalDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   20
         Top             =   2070
         Width           =   1695
      End
      Begin VB.Label lblJournalCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   300
         TabIndex        =   19
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   18
         Top             =   2490
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Journal As CJournal
Private m_ApArMass As Collection
Private m_ApAr As CAPARMas

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public ApArInd As Long

Private ApArText As String
Private FileName As String
Private m_MasterRef As CMasterRef

Public JournalType As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_Journal.SetFieldValue("JOURNAL_ID", ID)
      If Not glbDaily.QueryJournal(m_Journal, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Journal.PopulateFromRS(1, m_Rs)
      
      cboDepartment.ListIndex = IDToListIndex(cboDepartment, m_Journal.GetFieldValue("DEPARTMENT_ID"))
      txtJournalCode.Text = m_Journal.GetFieldValue("JOURNAL_NO")
      txtJournalDesc.Text = m_Journal.GetFieldValue("JOURNAL_DESC")
      cboApAr.ListIndex = IDToListIndex(cboApAr, m_Journal.GetFieldValue("APAR_IND"))
      uctlApArMasLookup.MyCombo.ListIndex = IDToListIndex(uctlApArMasLookup.MyCombo, m_Journal.GetFieldValue("APAR_MAS_ID"))
      chkPostFlag.Value = FlagToCheck(m_Journal.GetFieldValue("POST_FLAG"))
      uctlJournalDate.ShowDate = m_Journal.GetFieldValue("JOURNAL_DATE")
      
      CalculateSumDrCr
   Else
      ShowMode = SHOW_ADD
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("GL_JOURNAL_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("GL_JOURNAL_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblJournalCode, txtJournalCode, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblJournalDate, uctlJournalDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblApArMas, cboApAr, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(CUSTCODE_UNIQUE, txtJournalCode.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJournalCode.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_Journal.ShowMode = ShowMode
   Call m_Journal.SetFieldValue("JOURNAL_DATE", uctlJournalDate.ShowDate)
   Call m_Journal.SetFieldValue("POST_DATE", -1)
   Call m_Journal.SetFieldValue("POST_FLAG", Check2Flag(chkPostFlag.Value))
   Call m_Journal.SetFieldValue("DEPARTMENT_ID", cboDepartment.ItemData(Minus2Zero(cboDepartment.ListIndex)))
   Call m_Journal.SetFieldValue("JOURNAL_NO", txtJournalCode.Text)
   Call m_Journal.SetFieldValue("JOURNAL_DESC", txtJournalDesc.Text)
   Call m_Journal.SetFieldValue("APAR_MAS_ID", uctlApArMasLookup.MyCombo.ItemData(Minus2Zero(uctlApArMasLookup.MyCombo.ListIndex)))
   Call m_Journal.SetFieldValue("JOURNAL_TYPE", JournalType)
   Call m_Journal.SetFieldValue("JOURNAL_AMOUNT", 0)
         
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditJournal(m_Journal, IsOK, True, glbErrorLog) Then
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

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboApAr_Click()
Dim TempID As Long
   m_HasModify = True
   
   TempID = cboApAr.ItemData(Minus2Zero(cboApAr.ListIndex))
   If TempID <= 0 Then
      Exit Sub
   End If
   
   Call m_ApAr.SetFieldValue("APAR_MAS_ID", -1)
   Call m_ApAr.SetFieldValue("APAR_IND", TempID)
   Call LoadApArMas(m_ApAr, uctlApArMasLookup.MyCombo, m_ApArMass)
   Set uctlApArMasLookup.MyCollection = m_ApArMass
End Sub

Private Sub cboApAr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboDepartment_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboDepartment_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkPostFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkPostFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditJournalItem.ParentForm = Me
      Set frmAddEditJournalItem.TempCollection = m_Journal.JournalItems
      frmAddEditJournalItem.ShowMode = SHOW_ADD
      frmAddEditJournalItem.HeaderText = MapText("เพิ่มรายการสมุดรายวัน")
      Load frmAddEditJournalItem
      frmAddEditJournalItem.Show 1

      OKClick = frmAddEditJournalItem.OKClick

      Unload frmAddEditJournalItem
      Set frmAddEditJournalItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Journal.JournalItems)
         GridEX1.Rebind
         Call CalculateSumDrCr
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Set frmAddEditApArMasAccount.TempCollection = m_Journal.CstAccounts
'      frmAddEditApArMasAccount.ShowMode = SHOW_ADD
'      frmAddEditApArMasAccount.HeaderText = MapText("เพิ่มบัญชีลูกค้า")
'      Load frmAddEditApArMasAccount
'      frmAddEditApArMasAccount.Show 1
'
'      OKClick = frmAddEditApArMasAccount.OKClick
'
'      Unload frmAddEditApArMasAccount
'      Set frmAddEditApArMasAccount = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Journal.CstAccounts)
'         GridEX1.Rebind
'      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

'   If Trim(txtJournalCode.Text) = "" Then
'      Call glbDatabaseMngr.GenerateNumber(CUSTOMER_NUMBER, No, glbErrorLog)
'      txtJournalCode.Text = No
'   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Journal.JournalItems.Remove (ID2)
      Else
         m_Journal.JournalItems.Item(ID2).Flag = "D"
      End If

      GridEX1.ItemCount = CountItem(m_Journal.JournalItems)
      GridEX1.Rebind
      m_HasModify = True
      Call CalculateSumDrCr
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If m_Journal.CstAccounts.Item(ID2).MASTER_FLAG = "Y" Then
'         glbErrorLog.LocalErrorMsg = "ไม่สมารถลบบัญชีพื้นฐานได้"
'         glbErrorLog.ShowUserError
'         Exit Sub
'      End If
'
'      If ID1 <= 0 Then
'         m_Journal.CstAccounts.Remove (ID2)
'      Else
'         m_Journal.CstAccounts.Item(ID2).Flag = "D"
'      End If
'
'      GridEX1.itemcount = CountItem(m_Journal.CstAccounts)
'      GridEX1.Rebind
'      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

'   If Not VerifyAccessRight("GROUP_QUERY_RIGHT") Then
'      Exit Sub
'   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditJournalItem.ParentForm = Me
      frmAddEditJournalItem.ID = ID
      Set frmAddEditJournalItem.TempCollection = m_Journal.JournalItems
      frmAddEditJournalItem.HeaderText = MapText("แก้ไขรายการสมุดรายวัน")
      frmAddEditJournalItem.ShowMode = SHOW_EDIT
      Load frmAddEditJournalItem
      frmAddEditJournalItem.Show 1

      OKClick = frmAddEditJournalItem.OKClick

      Unload frmAddEditJournalItem
      Set frmAddEditJournalItem = Nothing

      If OKClick Then
         GridEX1.ItemCount = CountItem(m_Journal.JournalItems)
         GridEX1.Rebind
         Call CalculateSumDrCr
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      frmAddEditApArMasAccount.ID = ID
'      Set frmAddEditApArMasAccount.TempCollection = m_Journal.CstAccounts
'      frmAddEditApArMasAccount.HeaderText = MapText("แก้ไขบัญชีลูกค้า")
'      frmAddEditApArMasAccount.ShowMode = SHOW_EDIT
'      Load frmAddEditApArMasAccount
'      frmAddEditApArMasAccount.Show 1
'
'      OKClick = frmAddEditApArMasAccount.OKClick
'
'      Unload frmAddEditApArMasAccount
'      Set frmAddEditApArMasAccount = Nothing
'
'      If OKClick Then
'         GridEX1.itemcount = CountItem(m_Journal.CstAccounts)
'         GridEX1.Rebind
'      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
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
      ID = m_Journal.GetFieldValue("JOURNAL_ID")
      m_Journal.QueryFlag = 1
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

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call InitAPAR(cboApAr)
      
      Call m_MasterRef.SetFieldValue("MASTER_AREA", MASTER_DEPARTMENT)
      Call LoadMaster(m_MasterRef, cboDepartment)
            
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Journal.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Journal.QueryFlag = 0
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
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
   
   Set m_Journal = Nothing
   Set m_ApArMass = Nothing
   Set m_MasterRef = Nothing
   Set m_ApAr = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
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
   Col.Width = 1965
   Col.Caption = MapText("รหัสบัญชี")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 5100
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 2025
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("เดบิต")

   Set Col = GridEX1.Columns.add '6
   Col.Width = 2160
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("เครดิต")
End Sub

Private Sub InitGrid2()
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
   Col.Width = 1470
   Col.Caption = MapText("เลขที่บัญชี")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 6855
   Col.Caption = MapText("รายละเอียด")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 3240
   Col.Caption = MapText("แพคเกจ")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblJournalCode, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblJournalDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblDepartment, MapText("แผนก"))
   Call InitNormalLabel(lblJournalDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblApArMas, MapText("ลูกค้า/ผู้ค้า"))
   Call InitNormalLabel(lblDebit, MapText("เดบิต"))
   Call InitNormalLabel(lblCredit, MapText("เครดิต"))
   
   Call InitCombo(cboDepartment)
   Call InitCombo(cboApAr)
   
   Call txtJournalCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtJournalDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDebit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtDebit.Enabled = False
   Call txtCredit.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtCredit.Enabled = False
   
   Call InitCheckBox(chkPostFlag, "POST ไป GL")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("รายการ")
'   TabStrip1.Tabs.Add().Caption = MapText("บัญชีลูกค้า")
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
   Set m_Journal = New CJournal
   Set m_ApArMass = New Collection
   Set m_MasterRef = New CMasterRef
   Set m_ApAr = New CAPARMas
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Journal.JournalItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CJournalItem
      If m_Journal.JournalItems.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Journal.JournalItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.GetFieldValue("JOURNAL_ITEM_ID")
      Values(2) = RealIndex
      Values(3) = CR.GetFieldValue("ACC_CODE")
      Values(4) = CR.GetFieldValue("ITEM_DESC")
      If CR.GetFieldValue("DBCR_TYPE") = 1 Then
         Values(5) = FormatNumber(CR.GetFieldValue("DBCR_AMOUNT"))
         Values(6) = FormatNumber(0)
      ElseIf CR.GetFieldValue("DBCR_TYPE") = 2 Then
         Values(5) = FormatNumber(0)
         Values(6) = FormatNumber(CR.GetFieldValue("DBCR_AMOUNT"))
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If m_Journal.CstAccounts Is Nothing Then
'         Exit Sub
'      End If
'
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'
'      Dim Ca As CAccount
'      If m_Journal.CstAccounts.count <= 0 Then
'         Exit Sub
'      End If
'      Set Ca = GetItem(m_Journal.CstAccounts, RowIndex, RealIndex)
'      If Ca Is Nothing Then
'         Exit Sub
'      End If
'
'      Values(1) = Ca.ACCOUNT_ID
'      Values(2) = RealIndex
'      Values(3) = Ca.ACCOUNT_NO
'      Values(4) = Ca.NOTE
'      Values(5) = ""
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_Journal.JournalItems)
   GridEX1.Rebind
   
   Call CalculateSumDrCr
   m_HasModify = True
End Sub

Private Sub CalculateSumDrCr()
Dim Ji As CJournalItem
Dim SumDr As Double
Dim SumCr As Double

   SumDr = 0
   SumCr = 0
   For Each Ji In m_Journal.JournalItems
      If Ji.Flag <> "D" Then
         If Ji.GetFieldValue("DBCR_TYPE") = 1 Then
            SumDr = SumDr + Ji.GetFieldValue("DBCR_AMOUNT")
         ElseIf Ji.GetFieldValue("DBCR_TYPE") = 2 Then
            SumCr = SumCr + Ji.GetFieldValue("DBCR_AMOUNT")
         End If
      End If
   Next Ji
   txtDebit.Text = FormatNumber(SumDr)
   txtCredit.Text = FormatNumber(SumCr)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(m_Journal.JournalItems)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      Call InitGrid2
'      GridEX1.itemcount = CountItem(m_Journal.CstAccounts)
'      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtJournalDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtJournalCode_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub uctlJournalDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlApArMasLookup_Change()
   m_HasModify = True
End Sub
