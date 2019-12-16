VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddReceiptItem 
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   Icon            =   "frmAddReceiptItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   13875
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10440
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   18415
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   7380
         TabIndex        =   4
         Top             =   1230
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   7665
         Left            =   150
         TabIndex        =   6
         Top             =   1770
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   13520
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MultiSelect     =   -1  'True
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
         Column(1)       =   "frmAddReceiptItem.frx":27A2
         Column(2)       =   "frmAddReceiptItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddReceiptItem.frx":290E
         FormatStyle(2)  =   "frmAddReceiptItem.frx":2A6A
         FormatStyle(3)  =   "frmAddReceiptItem.frx":2B1A
         FormatStyle(4)  =   "frmAddReceiptItem.frx":2BCE
         FormatStyle(5)  =   "frmAddReceiptItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddReceiptItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   3
         Top             =   1230
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   7665
         Left            =   6900
         TabIndex        =   8
         Top             =   1770
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   13520
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
         Column(1)       =   "frmAddReceiptItem.frx":2F36
         Column(2)       =   "frmAddReceiptItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddReceiptItem.frx":30A2
         FormatStyle(2)  =   "frmAddReceiptItem.frx":31FE
         FormatStyle(3)  =   "frmAddReceiptItem.frx":32AE
         FormatStyle(4)  =   "frmAddReceiptItem.frx":3362
         FormatStyle(5)  =   "frmAddReceiptItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddReceiptItem.frx":34F2
      End
      Begin PorkShop.uctlTextBox txtKeySearch 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   770
         Width           =   1815
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtCustomer 
         Height          =   435
         Left            =   9405
         TabIndex        =   2
         Top             =   770
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextBox txtRef 
         Height          =   435
         Left            =   4860
         TabIndex        =   1
         Top             =   765
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin Threed.SSCheck chkShowAll 
         Height          =   405
         Left            =   13080
         TabIndex        =   18
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   714
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         TabIndex        =   17
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7680
         TabIndex        =   16
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   900
         Width           =   1635
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   6255
         TabIndex        =   7
         Top             =   5160
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   11340
         TabIndex        =   5
         Top             =   870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   14
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6090
         TabIndex        =   13
         Top             =   1320
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5400
         TabIndex        =   9
         Top             =   9660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReceiptItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7050
         TabIndex        =   10
         Top             =   9660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddReceiptItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_BillingDoc As CBillingDoc

Public CusID As Long
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Public DocumentType As SELL_BILLING_DOCTYPE

Private m_TempCol1 As Collection
Private m_TempCol2 As Collection
Private m_RetrunCnDn As Collection

Private Sub PopulateDestColl()
Dim Ri As CRcpCnDn_Item
Dim D As CBillingDoc

   For Each Ri In TempCollection
      Set D = New CBillingDoc
      
      If Ri.Flag <> "D" Then
         If Ri.Flag = "A" Then
            D.Flag = "A"
         ElseIf Ri.Flag = "I" Or Ri.Flag = "E" Then
            D.Flag = Ri.Flag
         End If
         D.BILLING_DOC_ID = Ri.GetFieldValue("DOC_ID")
         D.DOCUMENT_NO = Ri.GetFieldValue("DOC_NO")
         D.DOCUMENT_DATE = Ri.GetFieldValue("DOC_DATE")
         D.PAY_AMOUNT = Ri.GetFieldValue("ITEM_AMOUNT")
         D.PAID_AMOUNT = Ri.GetFieldValue("PAID_AMOUNT")
         D.DOCUMENT_TYPE = Ri.GetFieldValue("DOC_ID_TYPE")
         
         Call m_TempCol2.add(D)
      End If
      
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CBillingDoc
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.BILLING_DOC_ID = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function
Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim BD As CBillingDoc
Dim Temp As Double

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   
   While Not Rs.EOF
      Set BD = New CBillingDoc
      Call BD.PopulateFromRS(1009, Rs)
      
      If Not IsIn(m_TempCol2, BD.BILLING_DOC_ID) Then
         
         If BD.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
            BD.PAY_AMOUNT = BD.TOTAL_PRICE + BD.VAT_AMOUNT - (BD.DISCOUNT_AMOUNT + BD.EXT_DISCOUNT_AMOUNT + BD.PAID_AMOUNT)
         ElseIf BD.DOCUMENT_TYPE = RETURN_DOCTYPE Then
            BD.PAY_AMOUNT = BD.TOTAL_PRICE + BD.VAT_AMOUNT - (BD.DISCOUNT_AMOUNT + BD.EXT_DISCOUNT_AMOUNT + BD.PAID_AMOUNT)
         ElseIf BD.DOCUMENT_TYPE = CN_DOCTYPE Then 'จิวเพิ่มให้ใส่ VAT ด้วย ณ 04/10/2556
            If BD.PAID_AMOUNT > 0 Then
               BD.PAY_AMOUNT = 0
            Else
               BD.PAY_AMOUNT = BD.PAY_AMOUNT + BD.VAT_AMOUNT
            End If
         ElseIf BD.DOCUMENT_TYPE = DN_DOCTYPE Then
            If BD.PAID_AMOUNT > 0 Then
               BD.PAY_AMOUNT = 0
            Else
               BD.PAY_AMOUNT = BD.PAY_AMOUNT + BD.VAT_AMOUNT
            End If
         End If
         BD.PAID_AMOUNT = 0
         
         BD.Flag = ""
         If Check2Flag(chkShowAll.Value) <> "Y" Then
            If Val(FormatNumber(BD.PAY_AMOUNT, , False)) > 0 Then
               Call TempCol.add(BD, Trim(BD.DOCUMENT_NO))
               Call AddColl(m_RetrunCnDn, BD)
            End If
         Else
            Call TempCol.add(BD, Trim(BD.DOCUMENT_NO))
            Call AddColl(m_RetrunCnDn, BD)
         End If
      End If
      
      Set BD = Nothing
      Rs.MoveNext
   Wend
   
   
End Sub
Private Sub AddColl(Coll As Collection, BD As CBillingDoc)
On Error Resume Next
   'Call Coll.add(BD, Trim(Str(BD.REFER_TEXT)))
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      If Not VerifyTextControl(lblCustomer, txtCustomer, Not (txtCustomer.Enabled)) Then
         Exit Sub
      End If
      
      Call EnableForm(Me, False)
      
      m_BillingDoc.FROM_DATE = uctlFromDate.ShowDate
      m_BillingDoc.TO_DATE = uctlToDate.ShowDate
      m_BillingDoc.APAR_MAS_ID = CusID
      m_BillingDoc.APAR_CODE = PatchWildCard(txtCustomer.Text)
      m_BillingDoc.APAR_IND = 1
      m_BillingDoc.ORDER_TYPE = 1
      m_BillingDoc.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RETURN_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & ")"
      If DocumentType = RECEIPT2_DOCTYPE Or DocumentType = RECEIPT3_DOCTYPE Then
         m_BillingDoc.DOCUMENT_TYPE_RCP = RECEIPT2_DOCTYPE
      ElseIf DocumentType = BILLS_DOCTYPE Then
         m_BillingDoc.DOCUMENT_TYPE_RCP = BILLS_DOCTYPE
      End If
      m_BillingDoc.CANCEL_FLAG = "N"
      
      m_BillingDoc.SHOW_ALL_FLAG = Check2Flag(chkShowAll.Value)
      
      Call m_BillingDoc.QueryData(1009, m_Rs, ItemCount)
   End If
   
   If m_Rs.RecordCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

'   If Not m_HasModify Then
'      SaveData = True
'      Exit Function
'   End If
   
   Call PopulateTempColl
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub chkShowAll_Click(Value As Integer)
   If Not VerifyAccessRight("SECURE-ADMIN", "ระดับความปลอดภัย ADMIN ") Then
      chkShowAll.Value = ssCBUnchecked
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim BL As CBillingDoc
   
   If ID > 0 Then
      Set BL = TempCol1(ID)
      BL.Flag = ""
      BL.PAID_AMOUNT = BL.PAY_AMOUNT
      Call TempCol2.add(BL)
      TempCol1.Remove (ID)
   End If
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long
Dim BL As CBillingDoc
   For j = 1 To TempCol1.Count
      Set BL = TempCol1(j)
      BL.Flag = "A"
      BL.PAID_AMOUNT = BL.PAY_AMOUNT
      Call TempCol2.add(BL)
   Next j
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub
Private Sub cmdSelect_Click()
Dim TempID As Long
Dim Check As CBillingDoc
Dim ID As Long
Dim I As Long
Dim Row As Long
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   
   m_HasModify = True
   For Row = 1 To GridEX1.RowCount
      I = 0
      If GridEX1.RowSelected(Row) = True Then
         ID = GridEX1.GetRowData(Row).Value(1)
         For Each Check In m_TempCol1
            I = I + 1
            If Check.BILLING_DOC_ID = ID Then
               TempID = I
               Exit For
            End If
         Next Check
         Call CopyItem(m_TempCol1, m_TempCol2, TempID)
      End If
   Next Row

   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
      
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)
   
   GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub
Public Sub PopulateTempColl()
Dim D As CBillingDoc
Dim Ri As CRcpCnDn_Item
Dim I As Long

   For Each D In m_TempCol2
      I = I + 1
      If D.Flag = "" Then
         Set Ri = New CRcpCnDn_Item
         Ri.Flag = "A"
         Call Ri.SetFieldValue("DOC_ID", D.BILLING_DOC_ID)
         Call Ri.SetFieldValue("DOC_DATE", D.DOCUMENT_DATE)
         Call Ri.SetFieldValue("DOC_NO", D.DOCUMENT_NO)
         Call Ri.SetFieldValue("DOC_ID_TYPE", D.DOCUMENT_TYPE)
         
         If DocumentType = BILLS_DOCTYPE Then
            Call Ri.SetFieldValue("DOC_ID_BILLS", D.BILLING_DOC_ID)
         ElseIf DocumentType = RECEIPT2_DOCTYPE Or DocumentType = RECEIPT3_DOCTYPE Then
            Call Ri.SetFieldValue("DOC_ID_RCP", D.BILLING_DOC_ID)
         End If
         Call Ri.SetFieldValue("ITEM_AMOUNT", D.PAY_AMOUNT)
         Call Ri.SetFieldValue("PAID_AMOUNT", D.PAID_AMOUNT)
         
         Call TempCollection.add(Ri, Trim(Str(D.BILLING_DOC_ID)))
      ElseIf D.Flag = "I" Or D.Flag = "A" Or D.Flag = "E" Then
         Set Ri = GetObject("CRcpCnDn_Item", TempCollection, Trim(Str(D.BILLING_DOC_ID)), False)
         If Ri Is Nothing Then
            Set Ri = New CRcpCnDn_Item
            Ri.Flag = "A"
            Call TempCollection.add(Ri, Trim(Str(D.BILLING_DOC_ID)))
         End If
         If D.Flag = "I" Or D.Flag = "E" Then ' 1
            Ri.Flag = "E"
         End If
         Call Ri.SetFieldValue("DOC_ID", D.BILLING_DOC_ID)
         Call Ri.SetFieldValue("DOC_DATE", D.DOCUMENT_DATE)
         Call Ri.SetFieldValue("DOC_NO", D.DOCUMENT_NO)
         Call Ri.SetFieldValue("DOC_ID_TYPE", D.DOCUMENT_TYPE)
         
         If DocumentType = BILLS_DOCTYPE Then
            Call Ri.SetFieldValue("DOC_ID_BILLS", D.BILLING_DOC_ID)
         ElseIf DocumentType = RECEIPT2_DOCTYPE Or DocumentType = RECEIPT3_DOCTYPE Then
            Call Ri.SetFieldValue("DOC_ID_RCP", D.BILLING_DOC_ID)
         End If
         
         Call Ri.SetFieldValue("ITEM_AMOUNT", D.PAY_AMOUNT)
         Call Ri.SetFieldValue("PAID_AMOUNT", D.PAID_AMOUNT)
         
      End If
      
      Set Ri = Nothing
   Next D
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call PopulateDestColl
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_BillingDoc.QueryFlag = 1
      ElseIf ShowMode = SHOW_ADD Then
         m_BillingDoc.QueryFlag = 0
      End If
      
      If DocumentType = RECEIPT3_DOCTYPE Then
         uctlFromDate.ShowDate = DateAdd("YYYY", -1, Now)
         uctlToDate.ShowDate = Now
      Else
         uctlFromDate.ShowDate = DateAdd("YYYY", -1, Now)
         uctlToDate.ShowDate = Now
         txtCustomer.Enabled = False
         Call QueryData(True)
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
      Call cmdSearch_Click
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_BillingDoc = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set m_RetrunCnDn = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdSelect_Click
End Sub

Private Sub GridEX2_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
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
   Col.Width = 1425
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("หมายเลข")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 1500
   Col.Caption = MapText("ยอดค้าง")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("PO ลูกค้า")
   
   
   
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.ItemCount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX2.Columns.add '3
   Col.Width = 1300
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX2.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("หมายเลข")
   
   Set Col = GridEX2.Columns.add '5
   Col.Width = 1200
   Col.Caption = MapText("ส่วนลด")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX2.Columns.add '6
   Col.Width = 1200
   Col.Caption = MapText("รับชำระ")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = 1200
   Col.Caption = MapText("คงค้าง")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX2.Columns.add '8
   Col.Width = 0
   Col.Caption = MapText("ประเภทเอกสาร")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblDesc, MapText("หมายเลขเอกสาร"))
   Call InitNormalLabel(lblCustomer, MapText("รหัสลูกค้า"))
   Call InitNormalLabel(lblRef, MapText("อ้างอิง"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call txtCustomer.SetKeySearch("CUSTOMER_CODE")
   
   Call InitCheckBox(chkShowAll, "แสดงทั้งหมด")
    
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call txtCustomer.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdSelect, MapText(">"))
   
   Call InitGrid1
   Call InitGrid2
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
   Set m_BillingDoc = New CBillingDoc
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set m_RetrunCnDn = New Collection
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBillingDoc
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Values(5) = FormatNumber(CR.PAY_AMOUNT)
   Values(6) = CR.CUS_PO
   
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX2_DblClick()
Dim ID As Long

   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
            
   If Val(GridEX2.Value(8)) <> INVOICE_DOCTYPE Then
      Exit Sub
   End If
   
   ID = Val(GridEX2.Value(2))
   Set frmAddReceiptEditItem.ParentForm = Me
   Set frmAddReceiptEditItem.TempCollection = m_TempCol2
   frmAddReceiptEditItem.ID = ID
   frmAddReceiptEditItem.ShowMode = SHOW_EDIT
   frmAddReceiptEditItem.HeaderText = MapText("ใส่เงินรับชำระ")
   
   Load frmAddReceiptEditItem
   frmAddReceiptEditItem.Show 1

   OKClick = frmAddReceiptEditItem.OKClick

   Unload frmAddReceiptEditItem
   Set frmAddReceiptEditItem = Nothing

   If OKClick Then

      GridEX2.ItemCount = CountItem(m_TempCol2)
      GridEX2.Rebind
      
      m_HasModify = True
   End If
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CBillingDoc
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If
   
   Values(1) = CR.BILLING_DOC_ID
   Values(2) = RealIndex
   Values(3) = DateToStringExtEx2(CR.DOCUMENT_DATE)
   Values(4) = CR.DOCUMENT_NO
   Values(5) = 0
   Values(6) = FormatNumber(CR.PAID_AMOUNT)
   Values(7) = FormatNumber(CR.PAY_AMOUNT - CR.PAID_AMOUNT)
   Values(8) = CR.DOCUMENT_TYPE
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub RefreshGrid()

   GridEX2.ItemCount = CountItem(m_TempCol2)
   GridEX2.Rebind
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX2.Width = ScaleWidth - GridEX1.Left - GridEX1.Width - 1000
   GridEX2.Height = ScaleHeight - GridEX2.Top - cmdOK.Height - 100
   GridEX1.Height = GridEX2.Height
   cmdOK.Top = ScaleHeight - cmdOK.Height - 50
   cmdExit.Top = ScaleHeight - cmdExit.Height - 50
   cmdExit.Left = ScaleWidth / 2 + 50
   cmdOK.Left = ScaleWidth / 2 - cmdOK.Width - 50
End Sub
Private Sub txtKeySearch_LostFocus()
Dim BL As CBillingDoc
   If Len(Trim(txtKeySearch.Text)) <= 0 Then
      Exit Sub
   End If
   Set BL = GetObject("CBillingDoc", m_TempCol1, Trim(txtKeySearch.Text), False)
   If BL Is Nothing Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่มีเอกสารที่เลือก")
      glbErrorLog.ShowUserError
      txtKeySearch.SetFocus
      txtKeySearch.Text = ""
   Else
      BL.Flag = ""
      BL.PAID_AMOUNT = BL.PAY_AMOUNT
      Call m_TempCol2.add(BL)
      m_TempCol1.Remove (Trim(txtKeySearch.Text))
      txtKeySearch.SetFocus
      txtKeySearch.Text = ""
   End If
    GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Private Sub txtRef_LostFocus()
Dim BL As CBillingDoc
Dim BL1 As CBillingDoc
   
   If Len(Trim(txtRef.Text)) <= 0 Then
      Exit Sub
   End If
   Set BL1 = GetObject("CBillingDoc", m_RetrunCnDn, Trim(txtRef.Text), False)
   If BL1 Is Nothing Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่มีเอกสารที่เลือก")
      glbErrorLog.ShowUserError
      txtRef.SetFocus
      txtRef.Text = ""
      Exit Sub
   End If
   Set BL = GetObject("CBillingDoc", m_TempCol1, BL1.DOCUMENT_NO, False)
   
   Set BL1 = Nothing
   If BL Is Nothing Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่มีเอกสารที่เลือก")
      glbErrorLog.ShowUserError
      txtRef.SetFocus
      txtRef.Text = ""
   Else
      BL.Flag = ""
      BL.PAID_AMOUNT = BL.PAY_AMOUNT
      Call m_TempCol2.add(BL)
      m_TempCol1.Remove (Trim(BL.DOCUMENT_NO))
      txtRef.SetFocus
      txtRef.Text = ""
   End If
    GridEX1.ItemCount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.ItemCount = m_TempCol2.Count
   GridEX2.Rebind
End Sub
