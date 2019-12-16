VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditDoItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5550
   ClientLeft      =   4335
   ClientTop       =   240
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditDOItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   14340
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   5175
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   9128
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1350
         TabIndex        =   3
         Top             =   2520
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   1350
         TabIndex        =   0
         Top             =   1080
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtTotalPrice 
         Height          =   435
         Left            =   1365
         TabIndex        =   6
         Top             =   2970
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlProductLookup 
         Height          =   435
         Left            =   1365
         TabIndex        =   1
         Top             =   1560
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtDiscount 
         Height          =   435
         Left            =   8280
         TabIndex        =   8
         Top             =   3000
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtLeft 
         Height          =   435
         Left            =   8280
         TabIndex        =   9
         Top             =   3480
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   7680
         TabIndex        =   7
         Top             =   3000
         Width           =   555
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtAvgPrice 
         Height          =   435
         Left            =   8280
         TabIndex        =   5
         Top             =   2520
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlProductReturn 
         Height          =   435
         Left            =   1350
         TabIndex        =   2
         Top             =   2040
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtProDuctDetail 
         Height          =   435
         Left            =   6840
         TabIndex        =   28
         Top             =   1560
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdUnit 
         Height          =   435
         Left            =   3360
         TabIndex        =   4
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   767
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditDOItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblProductReturn 
         Alignment       =   1  'Right Justify
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   4005
         TabIndex        =   26
         Top             =   2640
         Width           =   1965
      End
      Begin VB.Label Label5 
         Height          =   345
         Left            =   10440
         TabIndex        =   25
         Top             =   3480
         Width           =   855
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   4920
         TabIndex        =   10
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label Label7 
         Height          =   345
         Left            =   10440
         TabIndex        =   24
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6960
         TabIndex        =   23
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Height          =   345
         Left            =   3405
         TabIndex        =   22
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   345
         Left            =   10440
         TabIndex        =   21
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6840
         TabIndex        =   20
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblProduct 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   165
         TabIndex        =   18
         Top             =   3090
         Width           =   1050
      End
      Begin VB.Label lblAvgPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6960
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   135
         TabIndex        =   16
         Top             =   1200
         Width           =   1125
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6600
         TabIndex        =   11
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8280
         TabIndex        =   12
         Top             =   3960
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   2610
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmAddEditDoItem"
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

Public ParentForm As Form
Public HeaderText As String
Public ID As Long
Public OKClick As Boolean

Public TempCollection As Collection

Public Area As Long

Private m_Products As Collection
Private m_ProductReturns As Collection
Private m_Locations As Collection

Public DocumentDate As Date
Public DocumentType As SELL_BILLING_DOCTYPE

Public CusID As Long

Private m_Sc As CStockCode
'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String
'--------------------------------------------------
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
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblQuantity, MapText("ปริมาณ"))
   Call InitNormalLabel(lblLocation, MapText("ที่จัดเก็บ"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคารวม"))
   Call InitNormalLabel(lblAvgPrice, MapText("ราคา/หน่วย"))
   Call InitNormalLabel(lblProduct, MapText("รายการ"))
   Call InitNormalLabel(lblProductReturn, MapText("รายการคืน"))
   Call InitNormalLabel(lblDiscount, MapText("ส่วนลด"))
   Call InitNormalLabel(Label4, MapText("บาท"))
   Call InitNormalLabel(lblLeft, MapText("คงค้าง"))
   Call InitNormalLabel(Label7, MapText("บาท"))
   Call InitNormalLabel(Label3, MapText("บาท"))
   Call InitNormalLabel(Label5, MapText("บาท"))
   Call InitNormalLabel(lblUnit, MapText(""))
   
   Call txtDiscountPercent.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtQuantity.SetTextLenType(TEXT_FLOAT, glbSetting.AMOUNT_LEN)
   Call txtAvgPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtTotalPrice.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtLeft.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtDiscount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.AMOUNT_LEN)
   Call txtProDuctDetail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   uctlProductLookup.MyTextBox.SetKeySearch ("STOCK_NO")
   
   txtLeft.Enabled = False
   
   lblProductReturn.Enabled = False
   uctlProductReturn.Enabled = False
      
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   cmdUnit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป (F7)"))
   Call InitMainButton(cmdUnit, MapText("U"))
   
End Sub
Private Sub CalculatePrice()
   txtLeft.Text = Val(txtTotalPrice.Text) - Val(txtDiscount.Text)
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim iCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CDocItem
         Set Di = TempCollection.Item(ID)
          uctlProductLookup.MyCombo.ListIndex = IDToListIndex(uctlProductLookup.MyCombo, Di.GetFieldValue("PART_ITEM_ID"))
          If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
            uctlProductReturn.MyCombo.ListIndex = IDToListIndex(uctlProductReturn.MyCombo, Di.GetFieldValue("PART_ITEM_RETURN_ID"))
          End If
          uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, Di.GetFieldValue("LOCATION_ID"))
          txtQuantity.Text = MyDiffEx(Di.GetFieldValue("ITEM_AMOUNT"), Di.GetFieldValue("UNIT_MULTIPLE"))
          txtAvgPrice.Text = Di.GetFieldValue("AVG_PRICE") * Di.GetFieldValue("UNIT_MULTIPLE")
          txtTotalPrice.Text = Di.GetFieldValue("TOTAL_PRICE")
          
         UnitID = Di.GetFieldValue("UNIT_TRAN_ID")
         Multiple = Di.GetFieldValue("UNIT_MULTIPLE")
         UnitName = Di.GetFieldValue("UNIT_TRAN_NAME")
          
         Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
          
          txtDiscountPercent.Text = Di.GetFieldValue("DISCOUNT_PERCENT")
          txtDiscount.Text = Di.GetFieldValue("DISCOUNT_AMOUNT")
          
          txtProDuctDetail.Text = Di.GetFieldValue("PRODUCT_DETAIL")
          
          Set Di = Nothing
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
      uctlProductLookup.MyCombo.ListIndex = -1
      txtQuantity.Text = ""
      txtAvgPrice.Text = ""
      txtDiscountPercent.Text = ""
      
   End If
   Call QueryData(True)
   Call ParentForm.RefreshGrid
      
   Call uctlProductLookup.SetFocus
End Sub


Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim TempID As Long
Dim SumAmount As Double
   
   If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      If Not VerifyCombo(lblProduct, uctlProductReturn.MyCombo, Not (uctlProductReturn.Enabled)) Then
         Exit Function
      End If
   End If
   
   If Not VerifyCombo(lblProduct, uctlProductLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, Not (uctlLocationLookup.Enabled)) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblTotalPrice, txtTotalPrice, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDiscount, txtDiscount, True) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If ShowMode = SHOW_ADD Then
'      If Not (LoadCheckBalance(Val(txtQuantity.Text) * Multiple, uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)), uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)), uctlProductLookup.MyTextBox.Text)) Then
'         Exit Function
'      End If
   End If
   
''------------------สำหรับตัดแบบ First In First out
'   Dim Part As CStockCode
'   Set Part = GetObject("CStockCode", m_Products, uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))
'
'   '-------------------------------------------------------------------------------------------------
'   If Part.LOT_FLAG = "Y" And LotItemLinkCollection.Count <= 0 And (DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE) Then
'      If Not GenerateAutoLotLink() Then
'         glbErrorLog.LocalErrorMsg = "ไม่มีจำนวน " & uctlProductLookup.MyCombo.Text & " เพียงพอสำหรับเบิก"
'         glbErrorLog.ShowUserError
'         'Exit Function
'      End If
'   ElseIf Part.LOT_FLAG = "Y" And LotItemLinkCollection.Count > 0 And (DocumentType = INVOICE_DOCTYPE Or DocumentType = RECEIPT1_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE) Then
'      If Not CheckLotItemAmount() Then
'         glbErrorLog.LocalErrorMsg = "จำนวน " & uctlProductLookup.MyCombo.Text & "ตัด LOT กับยอดเบิกไม่เท่ากัน กรุณาแก้ไขจำนวนทั้งคู่ให้เท่ากัน"
'         glbErrorLog.ShowUserError
'         Exit Function
'      End If
'    End If
''------------------สำหรับตัดแบบ First In First out
'Set Part = Nothing
'-------------------------------------------------------------------------------------------------
   
   Dim D As CAPARMas
   If Area = 2 Then
      Set D = m_SupplierColl(Trim(Str(CusID)))
   Else
      Set D = m_CustomerColl(Trim(Str(CusID)))
   End If
      
   Dim Di As CDocItem
   If ShowMode = SHOW_ADD Then
      Set Di = New CDocItem
      
      Di.Flag = "A"
      Call TempCollection.add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If
   
   If ShowMode = SHOW_EDIT Then
      
   End If
   
   Call Di.SetFieldValue("PART_ITEM_ID", uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex)))
   Call Di.SetFieldValue("STOCK_DESC", uctlProductLookup.MyCombo.Text)
   Call Di.SetFieldValue("STOCK_NO", uctlProductLookup.MyTextBox.Text)
   If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
      Call Di.SetFieldValue("PART_ITEM_RETURN_ID", uctlProductReturn.MyCombo.ItemData(Minus2Zero(uctlProductReturn.MyCombo.ListIndex)))
   End If
   
   Call Di.SetFieldValue("PRODUCT_DETAIL", txtProDuctDetail.Text)
   
   Call Di.SetFieldValue("LOCATION_ID", uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex)))
   Call Di.SetFieldValue("ITEM_AMOUNT", Val(txtQuantity.Text) * Multiple)
   Call Di.SetFieldValue("AVG_PRICE", MyDiffEx(Val(txtAvgPrice.Text), Multiple))
   Call Di.SetFieldValue("DISCOUNT_AMOUNT", Val(txtDiscount.Text))
   Call Di.SetFieldValue("DISCOUNT_PERCENT", Val(txtDiscountPercent.Text))
   Call Di.SetFieldValue("TOTAL_PRICE", FormatNumber(Val(txtQuantity.Text) * Val(txtAvgPrice.Text), , False))
   
   Call Di.SetFieldValue("UNIT_TRAN_ID", UnitID)
   Call Di.SetFieldValue("UNIT_MULTIPLE", Multiple)
   Call Di.SetFieldValue("UNIT_TRAN_NAME", UnitName)
   
   Call Di.SetFieldValue("ITEM_AMOUNT", Val(txtQuantity.Text) * Multiple)
   
   Set Di = Nothing
   SaveData = True
End Function
Private Sub cmdUnit_Click()
   frmChangeUnit.HeaderText = MapText("เปลี่ยนหน่วย")
   frmChangeUnit.UnitID = UnitID
   frmChangeUnit.Multiple = Multiple
   frmChangeUnit.UnitName = UnitName
   frmChangeUnit.UnitMName = UnitMName
   
   Load frmChangeUnit
   frmChangeUnit.Show 1
   
   UnitID = frmChangeUnit.UnitID
   Multiple = frmChangeUnit.Multiple
   UnitName = frmChangeUnit.UnitName
   UnitMName = frmChangeUnit.UnitMName
   
   Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   
   Unload frmChangeUnit
   Set frmChangeUnit = Nothing
   m_HasModify = True
   
   Call txtAvgPrice.SetFocus

End Sub

Private Sub Form_Activate()
Dim D As CAPARMas
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call LoadStockCode(uctlProductLookup.MyCombo, m_Products)
      Set uctlProductLookup.MyCollection = m_Products
      
      If Area = 2 Then
         Set D = m_SupplierColl(Trim(Str(CusID)))
      Else
         Set D = m_CustomerColl(Trim(Str(CusID)))
      End If
      
      
      If DocumentType = RETURN_DOCTYPE Or DocumentType = S_RETURN_DOCTYPE Then
         Call LoadStockCode(uctlProductReturn.MyCombo, m_ProductReturns)
         Set uctlProductReturn.MyCollection = m_ProductReturns
      End If
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
         
         '42 '002
         Dim Pt As CMasterRef
         For Each Pt In m_Locations
            If Left(Pt.KEY_CODE, 2) = "01" Then
               uctlLocationLookup.MyCombo.ListIndex = IDToListIndex(uctlLocationLookup.MyCombo, Pt.KEY_ID)
               Exit For
            End If
         Next Pt
         Set Pt = Nothing
         
      End If
      
      m_HasModify = False
      
      Call uctlProductLookup.SetFocus
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdNext_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
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
   Set m_Products = New Collection
   Set m_Locations = New Collection
   Set m_ProductReturns = New Collection
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_Products = Nothing
   Set m_Locations = Nothing
   Set m_Sc = Nothing
   Set m_ProductReturns = Nothing
   
End Sub

Private Sub txtAvgPrice_LostFocus()
   m_HasModify = True
    txtTotalPrice.Text = FormatNumber(Val(txtAvgPrice.Text) * Val(txtQuantity.Text), , False)
End Sub
Private Sub txtDiscount_Change()
   m_HasModify = True
   Call CalculatePrice
End Sub
Private Sub txtDiscountPercent_Change()
   m_HasModify = True
   txtDiscount.Text = Val(txtTotalPrice.Text) * Val(txtDiscountPercent.Text) / 100
End Sub
Private Sub txtLeft_Change()
   m_HasModify = True
End Sub
Private Sub txtProDuctDetail_Change()
   m_HasModify = True
End Sub

Private Sub txtQuantity_Change()
   m_HasModify = True
   txtTotalPrice.Text = FormatNumber(Val(txtAvgPrice.Text) * Val(txtQuantity.Text), , False)
End Sub
Private Sub txtTotalPrice_LostFocus()
Dim TotalPrice As Double
Dim Quantity As Double

   TotalPrice = Val(txtTotalPrice.Text)
   Quantity = Val(txtQuantity.Text)
   txtAvgPrice.Text = MyDiff(TotalPrice, Quantity)
   If Not (FormatNumber(TotalPrice - Val(txtDiscount.Text)) = FormatNumber(Val(txtAvgPrice.Text) * Quantity)) Then
      If ((TotalPrice - Val(txtDiscount.Text)) > (Val(txtAvgPrice.Text) * Quantity)) Then
         txtAvgPrice.Text = Val(txtAvgPrice.Text) + 0.01
         Call txtAvgPrice_LostFocus
         txtDiscount.Text = FormatNumber((Val(txtAvgPrice.Text) * Val(txtQuantity.Text)) - TotalPrice, , False)
      ElseIf (Val(txtTotalPrice.Text) - Val(txtDiscount.Text)) < (Val(txtAvgPrice.Text) * Val(txtQuantity.Text)) Then
         Call txtAvgPrice_LostFocus
         txtDiscount.Text = FormatNumber((Val(txtAvgPrice.Text) * Val(txtQuantity.Text)) - TotalPrice, , False)
      End If
   End If
   Call CalculatePrice
End Sub
Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlProductLookup_Change()
On Error Resume Next
Dim ID As Long
Dim D As CAPARMas
Dim PkgDetail As CPackageDetail

   ID = uctlProductLookup.MyCombo.ItemData(Minus2Zero(uctlProductLookup.MyCombo.ListIndex))
   If ID > 0 Then
      Set m_Sc = GetObject("CStockCode", m_Products, Trim(Str(ID)))
      
      UnitID = m_Sc.UNIT_ID
      Multiple = m_Sc.UNIT_AMOUNT
      UnitName = m_Sc.UNIT_NAME
      UnitMName = m_Sc.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
                        
      If Area = 1 Then
         Set D = m_CustomerColl(Trim(Str(CusID)))
      ElseIf Area = 2 Then
         Set D = m_SupplierColl(Trim(Str(CusID)))
      End If
         
      
      For Each PkgDetail In LoadPackageColl
         If D.PACKAGE_ID <= 0 Then
            If PkgDetail.GetFieldValue("PACKAGE_MASTER_FLAG") = "Y" And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID Then
               Exit For
            End If
         Else
            If PkgDetail.GetFieldValue("PACKAGE_ID") = D.PACKAGE_ID And PkgDetail.GetFieldValue("PART_ITEM_ID") = ID Then
               Exit For
            End If
         End If
      Next PkgDetail

      If Not (PkgDetail Is Nothing) Then
        If DocumentDate >= PkgDetail.GetFieldValue("PRO_FROM_DATE") And DocumentDate <= PkgDetail.GetFieldValue("PRO_TO_DATE") Then
           txtAvgPrice.Text = PkgDetail.GetFieldValue("PRO_ITEM_COST")
        Else
           txtAvgPrice.Text = PkgDetail.GetFieldValue("PART_ITEM_COST")
        End If
      Else
         txtAvgPrice.Text = 0
      End If
      
   Else
      lblUnit.Caption = ""
   End If
   
   
   m_HasModify = True

End Sub
Private Sub uctlProductReturn_Change()
   m_HasModify = True
End Sub
