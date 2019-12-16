VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTagetCDetail 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTagetCDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   6800
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Rebate.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlBranch 
         Height          =   435
         Left            =   2760
         TabIndex        =   0
         Top             =   120
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlCustomer 
         Height          =   435
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlEmployee 
         Height          =   435
         Left            =   2760
         TabIndex        =   2
         Top             =   1080
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   6480
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtTotalAmountRt 
         Height          =   435
         Left            =   3360
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtTotalPriceRt 
         Height          =   435
         Left            =   6480
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
         _ExtentX        =   9763
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   2760
         TabIndex        =   3
         Top             =   1560
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtVat 
         Height          =   435
         Left            =   2760
         TabIndex        =   6
         Top             =   2520
         Width           =   615
         _ExtentX        =   1931
         _ExtentY        =   767
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   1680
         Width           =   1965
      End
      Begin VB.Label lblTotalAmountRt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   2640
         Width           =   1725
      End
      Begin VB.Label lblTotalPriceRt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   2640
         Width           =   1725
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label lblEmployee 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1140
         Width           =   1845
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   1845
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2325
         TabIndex        =   9
         Top             =   3150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetCDetail.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3975
         TabIndex        =   10
         Top             =   3150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTagetCDetail.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5625
         TabIndex        =   11
         Top             =   3150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   2160
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmAddEditTagetCDetail"
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

Private m_TagetItem As CTagetDetail
Private m_TempEmp As CEmployee
Private m_Apm As CAPARMas
Private m_Branch As CMasterRef


Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public ParentForm As Form
Private BranchColl As Collection
Private EmployeeColl As Collection
Private m_Products As Collection

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
   
   Call InitNormalLabel(lblBranch, MapText("สาขา/เขตการขาย"))
   Call InitNormalLabel(lblCustomer, MapText("ลูกค้า"))
   Call InitNormalLabel(lblEmployee, MapText("พนักงานขาย"))
   Call InitNormalLabel(lblProduct, MapText("สินค้า"))
   
   Call InitNormalLabel(lblTotalAmount, MapText("จำนวนยอดขาย"))
   Call InitNormalLabel(lblTotalPrice, MapText("มูลค่ายอดขาย"))
   Call InitNormalLabel(lblTotalAmountRt, MapText("จำนวนยอดคืน"))
   Call InitNormalLabel(lblTotalPriceRt, MapText("มูลค่ายอดคืน"))
   
   Call txtTotalAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalAmountRt.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtTotalPriceRt.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   uctlCustomer.Enabled = False
   uctlEmployee.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim BD As CTagetDetail
         
         Set BD = TempCollection.Item(ID)
         
         uctlBranch.MyCombo.ListIndex = IDToListIndex(uctlBranch.MyCombo, BD.GetFieldValue("BRANCH_ID"))
         uctlCustomer.MyCombo.ListIndex = IDToListIndex(uctlCustomer.MyCombo, BD.GetFieldValue("APAR_MAS_ID"))
         uctlEmployee.MyCombo.ListIndex = IDToListIndex(uctlEmployee.MyCombo, BD.GetFieldValue("EMP_ID"))
         uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, BD.GetFieldValue("PART_ITEM_ID"))
         
         txtTotalAmount.Text = BD.GetFieldValue("TOTAL_AMOUNT")
         txtTotalPrice.Text = BD.GetFieldValue("TOTAL_PRICE")
         txtTotalAmountRt.Text = BD.GetFieldValue("TOTAL_AMOUNT_RT")
         txtTotalPriceRt.Text = BD.GetFieldValue("TOTAL_PRICE_RT")
         
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
       uctlBranch.MyCombo.ListIndex = -1
       uctlCustomer.MyCombo.ListIndex = -1
       uctlEmployee.MyCombo.ListIndex = -1
      txtTotalAmount.Text = ""
      txtTotalPrice.Text = ""
      txtTotalAmountRt.Text = ""
      txtTotalPriceRt.Text = ""
   End If
   Call QueryData(True)
   
   Call uctlBranch.SetFocus
   
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
Dim I As Long

   If Not VerifyCombo(lblBranch, uctlBranch.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblTotalAmount, txtTotalAmount, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim CheckBd As CTagetDetail
   For Each CheckBd In TempCollection
      I = I + 1
      
      If CheckBd.GetFieldValue("BRANCH_ID") = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)) And CheckBd.GetFieldValue("PART_ITEM_ID") = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)) And ID <> I Then
         glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & uctlBranch.MyCombo.Text & " และ " & uctlProductLookup.MyCombo.Text & " " & MapText("อยู่ในระบบแล้ว")
         glbErrorLog.ShowUserError
         Exit Function
      End If

   Next CheckBd
   
   Dim BD As CTagetDetail
   If ShowMode = SHOW_ADD Then
      Set BD = New CTagetDetail
      BD.Flag = "A"
      Call TempCollection.add(BD)
   Else
      Set BD = TempCollection.Item(ID)
      If BD.Flag <> "A" Then
         BD.Flag = "E"
      End If
   End If
   
   Call BD.SetFieldValue("BRANCH_ID", uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex)))
   Call BD.SetFieldValue("BRANCH_CODE", uctlBranch.MyTextBox.Text)
   Call BD.SetFieldValue("BRANCH_NAME", uctlBranch.MyCombo.Text)
   
   Call BD.SetFieldValue("APAR_MAS_ID", uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex)))
   Call BD.SetFieldValue("APAR_CODE", uctlCustomer.MyTextBox.Text)
   Call BD.SetFieldValue("APAR_NAME", uctlCustomer.MyCombo.Text)
   
   Call BD.SetFieldValue("EMP_ID", uctlEmployee.MyCombo.ItemData(Minus2Zero(uctlEmployee.MyCombo.ListIndex)))
   Call BD.SetFieldValue("EMPLOYEE_CODE", uctlEmployee.MyTextBox.Text)
   Call BD.SetFieldValue("EMPLOYEE_NAME", uctlEmployee.MyCombo.Text)
   
   Call BD.SetFieldValue("TOTAL_AMOUNT", Val(txtTotalAmount.Text))
   Call BD.SetFieldValue("TOTAL_PRICE", Val(txtTotalPrice.Text))
   Call BD.SetFieldValue("TOTAL_AMOUNT_RT", Val(txtTotalAmountRt.Text))
   Call BD.SetFieldValue("TOTAL_PRICE_RT", Val(txtTotalPriceRt.Text))
   
   Call BD.SetFieldValue("PART_ITEM_ID", uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))
   Call BD.SetFieldValue("STOCK_DESC", uctlProductLookup.MyCombo.Text)
   Call BD.SetFieldValue("STOCK_NO", uctlProductLookup.MyTextBox.Text)
   
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlBranch.MyCombo, BranchColl, , , MASTER_APARMAS_BRANCH)
      Set uctlBranch.MyCollection = BranchColl
      
      m_TempEmp.EMP_ID = -1
      Call LoadEmployee(m_TempEmp, uctlEmployee.MyCombo)
      Set uctlEmployee.MyCollection = m_EmployeeColl
      uctlEmployee.Visible = True
      
      Call LoadApArMas(m_Apm, uctlCustomer.MyCombo)
      Set uctlCustomer.MyCollection = m_CustomerColl
      uctlCustomer.Visible = True
      
      Call LoadStockCode(uctlProductLookup.MyCombo, m_Products)
      Set uctlProductLookup.MyCollection = m_Products
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
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

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_TagetItem = New CTagetDetail
   Set m_TempEmp = New CEmployee
   Set m_Apm = New CAPARMas
   Set m_Branch = New CMasterRef
   
   Set BranchColl = New Collection
   Set EmployeeColl = New Collection
   Set m_Products = New Collection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_TagetItem = Nothing
   Set m_TempEmp = Nothing
   Set m_Apm = Nothing
   Set m_Branch = Nothing
   
   Set BranchColl = Nothing
   Set EmployeeColl = Nothing
   Set m_Products = Nothing
End Sub

Private Sub txtTotalAmount_Change()
Dim PkgDetail As CPackageDetail
Dim D As CAPARMas
Dim ID As Long
Dim Pi As CStockCode
   If m_HasActivate Then
      ID = uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex))
      If ID > 0 Then
         Set D = m_CustomerColl(Trim(Str(uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex)))))
         
         For Each PkgDetail In LoadPackageColl
            If D.PACKAGE_ID <= 0 Then
               If PkgDetail.GetFieldValue("PACKAGE_MASTER_FLAG") = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)) Then
                  Exit For
               End If
            Else
               If PkgDetail.GetFieldValue("PACKAGE_ID") = D.PACKAGE_ID And PkgDetail.GetFieldValue("PART_ITEM_ID") = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)) Then
                  Exit For
               End If
            End If
         Next PkgDetail
         
         If Not (PkgDetail Is Nothing) Then
            Set Pi = GetObject("CStockCode", m_Products, Trim(Str(uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))))
            txtTotalPrice.Text = MyDiffEx(Val(txtTotalAmount.Text) * PkgDetail.GetFieldValue("PART_ITEM_COST"), Pi.UNIT_AMOUNT)
         Else
            txtTotalPrice.Text = ""
         End If
      End If
   End If
End Sub
Private Sub txtTotalAmountRt_Change()
Dim PkgDetail As CPackageDetail
Dim D As CAPARMas
Dim ID As Long
Dim Pi As CStockCode
   If m_HasActivate Then
      ID = uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex))
      If ID > 0 Then
         Set D = m_CustomerColl(Trim(Str(uctlCustomer.MyCombo.ItemData(Minus2Zero(uctlCustomer.MyCombo.ListIndex)))))
         
         For Each PkgDetail In LoadPackageColl
            If D.PACKAGE_ID <= 0 Then
               If PkgDetail.GetFieldValue("PACKAGE_MASTER_FLAG") = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)) Then
                  Exit For
               End If
            Else
               If PkgDetail.GetFieldValue("PACKAGE_ID") = D.PACKAGE_ID And PkgDetail.GetFieldValue("PART_ITEM_ID") = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)) Then
                  Exit For
               End If
            End If
         Next PkgDetail
         
         If Not (PkgDetail Is Nothing) Then
            Set Pi = GetObject("CStockCode", m_Products, Trim(Str(uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))))
            txtTotalPriceRt.Text = MyDiffEx(Val(txtTotalAmountRt.Text) * PkgDetail.GetFieldValue("PART_ITEM_COST"), Pi.UNIT_AMOUNT)
         Else
            txtTotalPriceRt.Text = ""
         End If
      End If
   End If
End Sub
Private Sub txtTotalPrice_Change()
   m_HasModify = True
End Sub
Private Sub txtTotalPriceRt_Change()
   m_HasModify = True
End Sub

Private Sub txtVat_LostFocus()
   txtTotalAmountRt.Text = Val(txtTotalAmount.Text) * Val(txtVat.Text) / 100
   txtTotalPriceRt.Text = Val(txtTotalPrice.Text) * Val(txtVat.Text) / 100
End Sub

Private Sub uctlBranch_Change()
Dim ID As Long
Dim Ba As CMasterRef
   ID = uctlBranch.MyCombo.ItemData(Minus2Zero(uctlBranch.MyCombo.ListIndex))
   If ID > 0 Then
      Set Ba = BranchColl(Trim(Str(ID)))
      uctlCustomer.MyCombo.ListIndex = IDToListIndex(uctlCustomer.MyCombo, Ba.PARENT_EX_ID2)
      uctlEmployee.MyCombo.ListIndex = IDToListIndex(uctlEmployee.MyCombo, Ba.PARENT_EX_ID)
   End If
   
   m_HasModify = True
End Sub

Private Sub uctlCustomer_Change()
   m_HasModify = True
End Sub

Private Sub uctlEmployee_Change()
   m_HasModify = True
End Sub
