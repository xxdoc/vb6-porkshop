VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditCashDoc 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditCashDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextLookup uctlBankLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1830
         Width           =   8300
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6300
         TabIndex        =   2
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   8
         Top             =   3680
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
      Begin PorkShop.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   900
         Width           =   2385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3495
         Left            =   150
         TabIndex        =   9
         Top             =   4230
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   6165
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
         Column(1)       =   "frmAddEditCashDoc.frx":27A2
         Column(2)       =   "frmAddEditCashDoc.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditCashDoc.frx":290E
         FormatStyle(2)  =   "frmAddEditCashDoc.frx":2A6A
         FormatStyle(3)  =   "frmAddEditCashDoc.frx":2B1A
         FormatStyle(4)  =   "frmAddEditCashDoc.frx":2BCE
         FormatStyle(5)  =   "frmAddEditCashDoc.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditCashDoc.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextLookup uctlBankBranchLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   2280
         Width           =   8300
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlBankAccountLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1380
         Width           =   8300
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlEmployeeLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   2700
         Width           =   8300
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlCustomerLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   25
         Top             =   3150
         Width           =   8300
         _ExtentX        =   14631
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   24
         Top             =   2730
         Width           =   1455
      End
      Begin VB.Label lblBankAccount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   330
         TabIndex        =   23
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lblBankBranch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   210
         TabIndex        =   22
         Top             =   2340
         Width           =   1455
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4260
         TabIndex        =   1
         Top             =   900
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCheck chkCommit 
         Height          =   435
         Left            =   10320
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblBank 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   210
         TabIndex        =   21
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   20
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   19
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   14
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   15
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":3884
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCashDoc.frx":3B9E
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditCashDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CashDoc As CCashDoc

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As CASH_DOC_TYPE

Private m_Employee As CEmployee
Private m_Mr As CMasterRef

Private Mr As CMasterRef

Private m_Banks As Collection
Private m_BankBranchs As Collection
Private m_BankAccounts As Collection
Private m_Employees As Collection
Private m_Customers  As Collection
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_CashDoc.SetFieldValue("CASH_DOC_ID", ID)
      m_CashDoc.QueryFlag = 1
      If Not glbDaily.QueryCashDoc(m_CashDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_CashDoc.PopulateFromRS(1, m_Rs)
      
      uctlDocumentDate.ShowDate = m_CashDoc.GetFieldValue("DOCUMENT_DATE")
      txtDocumentNo.Text = m_CashDoc.GetFieldValue("DOCUMENT_NO")
      uctlBankLookup.MyCombo.ListIndex = IDToListIndex(uctlBankLookup.MyCombo, m_CashDoc.GetFieldValue("BANK_ID"))
      uctlBankBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBankBranchLookup.MyCombo, m_CashDoc.GetFieldValue("BANK_BRANCH"))
      uctlBankAccountLookup.MyCombo.ListIndex = IDToListIndex(uctlBankAccountLookup.MyCombo, m_CashDoc.GetFieldValue("BANK_ACCOUNT"))
      uctlEmployeeLookUp.MyCombo.ListIndex = IDToListIndex(uctlEmployeeLookUp.MyCombo, m_CashDoc.GetFieldValue("EMP_ID"))
      uctlCustomerLookUp.MyCombo.ListIndex = IDToListIndex(uctlCustomerLookUp.MyCombo, m_CashDoc.GetFieldValue("APAR_MAS_ID"))
      
      If DocumentType = CASH_DEPOSIT Then
         Call glbDaily.CreateCashTransferItems(m_CashDoc)
      End If
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBank, uctlBankLookup.MyCombo, Not uctlBankLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankBranch, uctlBankBranchLookup.MyCombo, Not uctlBankBranchLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBankAccount, uctlBankAccountLookup.MyCombo, Not uctlBankAccountLookup.Enabled) Then
      Exit Function
   End If
   If Not VerifyCombo(lblEmployee, uctlEmployeeLookUp.MyCombo, Not uctlEmployeeLookUp.Enabled) Then
      Exit Function
   End If
'   If Not (DocumentType = POST_CHEQUE Or DocumentType = WAITING_CHEQUE Or DocumentType = PASSED_CHEQUE) Then
'      If Not VerifyCombo(lblCustomer, uctlCustomerLookup.MyCombo, False) Then
'         Exit Function
'      End If
'   End If
   
'   If Not glbDaily.VerifyDrCr(m_CashDoc.JournalItems) Then
'      glbErrorLog.LocalErrorMsg = "ผลรวมของ เดบิต จะต้องเท่ากับผลรวมของ เครดิต"
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If

'   If Not CheckUniqueNs(EXPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_CashDoc.ShowMode = ShowMode
   Call m_CashDoc.SetFieldValue("CASH_DOC_ID", ID)
    Call m_CashDoc.SetFieldValue("DOCUMENT_DATE", uctlDocumentDate.ShowDate)
   Call m_CashDoc.SetFieldValue("DOCUMENT_NO", txtDocumentNo.Text)
   Call m_CashDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call m_CashDoc.SetFieldValue("BANK_ID", uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("BANK_BRANCH", uctlBankBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBankBranchLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("BANK_ACCOUNT", uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("EMP_ID", uctlEmployeeLookUp.MyCombo.ItemData(Minus2Zero(uctlEmployeeLookUp.MyCombo.ListIndex)))
   Call m_CashDoc.SetFieldValue("APAR_MAS_ID", uctlCustomerLookUp.MyCombo.ItemData(Minus2Zero(uctlCustomerLookUp.MyCombo.ListIndex)))
   
   
   Call EnableForm(Me, False)
   If (DocumentType = CASH_DEPOSIT) Then
      Call CreateCashTranItems
   End If
   
   If Not glbDaily.AddEditCashDoc(m_CashDoc, IsOK, True, glbErrorLog) Then
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
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub
Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim oMenu As CPopupMenu
Dim lMenuChoosen As Long

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If DocumentType = CASH_DEPOSIT Then
         Set oMenu = New CPopupMenu
         lMenuChoosen = oMenu.Popup("เงินสดในมือ", "-", "เช็คในมือ")
         Set oMenu = Nothing

         If lMenuChoosen = 1 Then
            frmAddEditCashTran5.DocumentType = DocumentType
            Set frmAddEditCashTran5.ParentForm = Me
            Set frmAddEditCashTran5.TempCollection = m_CashDoc.TransferItems
            frmAddEditCashTran5.ShowMode = SHOW_ADD
            frmAddEditCashTran5.HeaderText = MapText("เพิ่ม" & "รายการนำฝากเงิน")
            Load frmAddEditCashTran5
            frmAddEditCashTran5.Show 1

            OKClick = frmAddEditCashTran5.OKClick

            Unload frmAddEditCashTran5
            Set frmAddEditCashTran5 = Nothing
         ElseIf lMenuChoosen = 3 Then

            If Not VerifyCombo(lblCustomer, uctlCustomerLookUp.MyCombo, False) Then
               Exit Sub
            End If

            frmAddChequeItem.DocumentType = DocumentType
            frmAddChequeItem.ApArID = uctlCustomerLookUp.MyCombo.ItemData(Minus2Zero(uctlCustomerLookUp.MyCombo.ListIndex))
            Set frmAddChequeItem.TempCollection = m_CashDoc.TransferItems
            frmAddChequeItem.ShowMode = SHOW_ADD
            frmAddChequeItem.HeaderText = MapText("เพิ่ม" & "รายการนำฝากเช็ค")
            Load frmAddChequeItem
            frmAddChequeItem.Show 1

            OKClick = frmAddChequeItem.OKClick

            Unload frmAddChequeItem
            Set frmAddChequeItem = Nothing
         End If

         If OKClick Then
            m_HasModify = True
            GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
            GridEX1.Rebind
         End If
      ElseIf DocumentType = POST_CHEQUE Then
         If Not VerifyCombo(lblBankAccount, uctlBankAccountLookup.MyCombo, False) Then
            Exit Sub
         End If

         frmAddChequeItemEx.DocumentType = DocumentType
         frmAddChequeItemEx.PostType = POST_CLEAR
         frmAddChequeItemEx.AccountID = uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex))
         Set frmAddChequeItemEx.TempCollection = m_CashDoc.PostItems
         frmAddChequeItemEx.ShowMode = SHOW_ADD
         frmAddChequeItemEx.HeaderText = MapText("เพิ่ม" & "รายการเช็คที่ชึ้นเงินแล้ว")
         Load frmAddChequeItemEx
         frmAddChequeItemEx.Show 1

         OKClick = frmAddChequeItemEx.OKClick

         Unload frmAddChequeItemEx
         Set frmAddChequeItemEx = Nothing

         If OKClick Then
            m_HasModify = True
            GridEX1.ItemCount = CountItem(m_CashDoc.PostItems)
            GridEX1.Rebind
         End If
      End If
   End If
   
End Sub
Private Sub CreateCashTranItems()
Dim Ti As CCashTransferItem
Dim Ei As CCashTran
Dim II As CCashTran

   Set m_CashDoc.CashTranItems = Nothing
   Set m_CashDoc.CashTranItems = New Collection

   For Each Ti In m_CashDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set II = Ti.ImportItem

      Ei.Flag = Ti.Flag
      II.Flag = Ti.Flag

      Call m_CashDoc.CashTranItems.add(Ei)
      Call m_CashDoc.CashTranItems.add(II)
   Next Ti
End Sub
Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If DocumentType = CASH_DEPOSIT Then
         If ID1 <= 0 Then
            m_CashDoc.TransferItems.Remove (ID2)
         Else
            m_CashDoc.TransferItems.Item(ID2).Flag = "D"
         End If
      ElseIf DocumentType = POST_CHEQUE Then
         If ID1 <= 0 Then
            m_CashDoc.PostItems.Remove (ID2)
         Else
            m_CashDoc.PostItems.Item(ID2).Flag = "D"
         End If
      End If
     Call RefreshGrid(DocumentType, True)
   End If
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
Dim PaymentType As Long
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
    If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
        If (DocumentType = POST_CHEQUE) Then
            glbErrorLog.LocalErrorMsg = "ไม่สามารถที่จะเปิดมาแก้ไขได้กรุณาลบแล้วสร้างใหม่"
            glbErrorLog.ShowUserError
            Exit Sub
        ElseIf (DocumentType = CASH_DEPOSIT) Then
            PaymentType = GridEX1.Value(8)
            If PaymentType = CHEQUE_HAND_PMT Or PaymentType = CHEQUE_BANK_PMT Then   'เช็ค
                glbErrorLog.LocalErrorMsg = "รายการนำฝากเช็คไม่สามารถที่จะเปิดมาแก้ไขได้"
                glbErrorLog.ShowUserError
                Exit Sub
            End If
            
            frmAddEditCashTran5.DocumentType = DocumentType
            Set frmAddEditCashTran5.ParentForm = Me
            frmAddEditCashTran5.ID = ID
            Set frmAddEditCashTran5.TempCollection = m_CashDoc.TransferItems
            frmAddEditCashTran5.HeaderText = MapText("แก้ไข" & "รายการนำฝากเงิน")
            frmAddEditCashTran5.ShowMode = SHOW_EDIT
            Load frmAddEditCashTran5
            frmAddEditCashTran5.Show 1
            
            OKClick = frmAddEditCashTran5.OKClick
            
            Unload frmAddEditCashTran5
            Set frmAddEditCashTran5 = Nothing
            
            If OKClick Then
                m_HasModify = True
                GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
                GridEX1.Rebind
            End If
        End If
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
      ID = m_CashDoc.GetFieldValue("CASH_DOC_ID")
      m_CashDoc.QueryFlag = 1
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
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Dim ApAr As CAPARMas
      Set ApAr = New CAPARMas
      Call LoadApArMas(ApAr, uctlCustomerLookUp.MyCombo)
      Set uctlCustomerLookUp.MyCollection = m_CustomerColl
      Set ApAr = Nothing
      
      Call LoadMaster(uctlBankLookup.MyCombo, m_Banks, , , MASTER_BANK)
      Set uctlBankLookup.MyCollection = m_Banks
      
      Call LoadMaster(uctlBankBranchLookup.MyCombo, m_BankBranchs, , , MASTER_BBRANCH)
      Set uctlBankBranchLookup.MyCollection = m_BankBranchs
      
      Call LoadMaster(uctlBankAccountLookup.MyCombo, m_BankAccounts, , , MASTER_BANK_ACCOUNT)
      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
      
      Dim Emp As CEmployee
      Set Emp = New CEmployee
      Call LoadEmployee(Emp, uctlEmployeeLookUp.MyCombo)
      Set uctlEmployeeLookUp.MyCollection = m_EmployeeColl
      Set Emp = Nothing
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_CashDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         uctlDocumentDate.ShowDate = Now
         m_CashDoc.QueryFlag = 0
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
   ElseIf Shift = 0 And KeyCode = 123 Then
'      Call AddMemoNote
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_CashDoc = Nothing
   Set m_Employees = Nothing
   Set m_Employee = Nothing
   Set m_Banks = Nothing
   Set m_BankBranchs = Nothing
   Set m_BankAccounts = Nothing
   Set m_Employee = Nothing
   Set m_Customers = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1(Ind As CASH_DOC_TYPE)
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

   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If (Ind = CASH_DEPOSIT) Then
         Set Col = GridEX1.Columns.add '3
         Col.Width = 2400
         Col.Caption = MapText("ประเภท")

         Set Col = GridEX1.Columns.add '4
         Col.Width = 2415
         Col.TextAlignment = jgexAlignRight
         Col.Caption = MapText("จำนวนเงิน")

         Set Col = GridEX1.Columns.add '5
         Col.Width = 2745
         Col.Caption = MapText("เลขที่เช็ค")

         Set Col = GridEX1.Columns.add '6
         Col.Width = 2820
         Col.Caption = MapText("วันที่เช็ค")

         Set Col = GridEX1.Columns.add '7
         Col.Width = 3570
         Col.Caption = MapText("วันที่ขึ้นเงิน")

         Set Col = GridEX1.Columns.add '8
         Col.Width = 0
         Col.Visible = False
         Col.Caption = MapText("PAYMENT_TYPE")
         
      ElseIf (Ind = POST_CHEQUE) Then

         Set Col = GridEX1.Columns.add '3
         Col.Width = 1500
         Col.Caption = MapText("เลขที่เช็ค")

         Set Col = GridEX1.Columns.add '5
         Col.Width = 1500
         Col.Caption = MapText("วันที่เช็ค")

         Set Col = GridEX1.Columns.add '4
         Col.Width = 1500
         Col.TextAlignment = jgexAlignRight
         Col.Caption = MapText("จำนวนเงิน")

         Set Col = GridEX1.Columns.add '5
         Col.Width = 2820
         Col.Caption = MapText("ธนาคาร")

         Set Col = GridEX1.Columns.add '6
         Col.Width = 3570
         Col.Caption = MapText("สาขาธนาคาร")
      End If
   End If
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblBank, MapText("ธนาคาร"))
   Call InitNormalLabel(lblBankBranch, MapText("สาขาธนาคาร"))
   Call InitNormalLabel(lblBankAccount, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
   
   If DocumentType = POST_CHEQUE Then
      Call InitNormalLabel(lblEmployee, MapText("ผู้ตรวจสอบ"))
      uctlCustomerLookUp.Enabled = False
      Label4.Visible = False
   Else
      Call InitNormalLabel(lblEmployee, MapText("พนักงาน"))
   End If
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
    
    uctlBankBranchLookup.Enabled = False
    uctlBankLookup.Enabled = False
    
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))

   Call InitCheckBox(chkCommit, MapText("ห้ามแก้ไข"))

   Call InitGrid1(DocumentType)

   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16

   Dim T As Object
   TabStrip1.Tabs.Clear

   If DocumentType = CASH_DEPOSIT Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการนำฝาก")
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = POST_CHEQUE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการเช็คที่ขึ้นเงินได้แล้ว")
      T.Tag = DocumentType & "-1"
   End If

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
   Set m_CashDoc = New CCashDoc
   Set m_Employee = New CEmployee
   Set m_Employees = New Collection
   Set m_Customers = New Collection
   Set m_Mr = New CMasterRef
   Set m_Banks = New Collection
   Set m_BankBranchs = New Collection
   Set m_BankAccounts = New Collection
   Set m_Employee = New CEmployee
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
Dim TR As CCashTransferItem
Dim Ct1 As CCashTran
Dim Pos As CCashDocPost

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If m_CashDoc.CashTranItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

        If DocumentType = POST_CHEQUE Then
         If m_CashDoc.PostItems.Count <= 0 Then
            Exit Sub
         End If
         Set Pos = GetItem(m_CashDoc.PostItems, RowIndex, RealIndex)
         If Pos Is Nothing Then
            Exit Sub
         End If

         Values(1) = Pos.GetFieldValue("CASH_DOC_POST_ID")
         Values(2) = RealIndex
         Values(3) = Pos.GetFieldValue("CHEQUE_NO")
         Values(4) = DateToStringExtEx2(Pos.GetFieldValue("CHEQUE_DATE"))
         Values(5) = FormatNumber(Pos.GetFieldValue("CHEQUE_AMOUNT"))
         Values(6) = Pos.GetFieldValue("BANK_NAME")
         Values(7) = Pos.GetFieldValue("BRANCH_NAME")
    ElseIf DocumentType = CASH_DEPOSIT Then
         If m_CashDoc.TransferItems.Count <= 0 Then
            Exit Sub
         End If
         Set TR = GetItem(m_CashDoc.TransferItems, RowIndex, RealIndex)
         If TR Is Nothing Then
            Exit Sub
         End If

         Values(1) = TR.ImportItem.GetFieldValue("CASH_TRAN_ID")
         Values(2) = RealIndex
         Values(3) = TR.ExportItem.GetFieldValue("PAYMENT_TYPE_NAME")
         Values(4) = FormatNumber(TR.ImportItem.GetFieldValue("AMOUNT"))
         Values(5) = TR.ExportItem.Cheque.GetFieldValue("CHEQUE_NO")
         Values(6) = DateToStringExtEx2(TR.ExportItem.Cheque.GetFieldValue("CHEQUE_DATE"))
         Values(7) = DateToStringExtEx2(TR.ExportItem.Cheque.GetFieldValue("EFFECTIVE_DATE"))
         Values(8) = TR.ExportItem.GetFieldValue("PAYMENT_TYPE")
    End If
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   Call InitGrid1(DocumentType)
   Call RefreshGrid(DocumentType, False)
End Sub
Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub
Private Sub uctlBankLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlBankAccountLookup_Change()
Dim TempID1 As Long
   
   TempID1 = uctlBankAccountLookup.MyCombo.ItemData(Minus2Zero(uctlBankAccountLookup.MyCombo.ListIndex))
   If TempID1 > 0 Then
      Set Mr = GetObject("CMasterRef", m_BankAccounts, Trim(Str(TempID1)))
      uctlBankLookup.MyCombo.ListIndex = IDToListIndex(uctlBankLookup.MyCombo, Mr.PARENT_EX_ID4)
      uctlBankBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBankBranchLookup.MyCombo, Mr.PARENT_EX_ID5)
   Else
      uctlBankLookup.MyCombo.ListIndex = -1
      uctlBankBranchLookup.MyCombo.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlBankBranchLookup_Change()
Dim TempID1 As Long
Dim TempID2 As Long
   
   TempID1 = uctlBankLookup.MyCombo.ItemData(Minus2Zero(uctlBankLookup.MyCombo.ListIndex))
   TempID2 = uctlBankBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBankBranchLookup.MyCombo.ListIndex))
   
   If TempID2 > 0 Then
'      Call LoadMaster(uctlBankAccountLookup.MyCombo, m_BankAccounts, BANK_ACCOUNT, TempID1, TempID2)
'      Set uctlBankAccountLookup.MyCollection = m_BankAccounts
   End If
   
   m_HasModify = True
End Sub
Public Sub RefreshGrid(Ind As CASH_DOC_TYPE, Flag As Boolean)
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If Ind = CASH_DEPOSIT Then
         GridEX1.ItemCount = CountItem(m_CashDoc.TransferItems)
         GridEX1.Rebind
      ElseIf (Ind = POST_CHEQUE) Then
         GridEX1.ItemCount = CountItem(m_CashDoc.PostItems)
         GridEX1.Rebind
      End If
   End If
   
   If Flag Then
      m_HasModify = Flag
   End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
