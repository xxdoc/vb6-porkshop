VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBarcodeProduction 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmBarcodeProduction.frx":0000
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
      TabIndex        =   14
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Rebate.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   5340
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   870
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   1800
         TabIndex        =   8
         Top             =   2640
         Width           =   10035
         _ExtentX        =   17701
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
      Begin Rebate.uctlTextBox txtDocumentNo 
         Height          =   435
         Left            =   2250
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   1785
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
         Height          =   1365
         Left            =   150
         TabIndex        =   9
         Top             =   3120
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   2408
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
         Column(1)       =   "frmBarcodeProduction.frx":27A2
         Column(2)       =   "frmBarcodeProduction.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBarcodeProduction.frx":290E
         FormatStyle(2)  =   "frmBarcodeProduction.frx":2A6A
         FormatStyle(3)  =   "frmBarcodeProduction.frx":2B1A
         FormatStyle(4)  =   "frmBarcodeProduction.frx":2BCE
         FormatStyle(5)  =   "frmBarcodeProduction.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmBarcodeProduction.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtTotalAmount 
         Height          =   435
         Left            =   10260
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   870
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtQuantity 
         Height          =   435
         Left            =   1710
         TabIndex        =   1
         Top             =   2010
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlPartLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   0
         Top             =   1440
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtTotalPrice 
         Height          =   465
         Left            =   7965
         TabIndex        =   2
         Top             =   2010
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   555
         Left            =   1800
         TabIndex        =   23
         Top             =   5880
         Width           =   9915
         _ExtentX        =   17489
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
      Begin Rebate.uctlTextBox txtQuantityIN 
         Height          =   435
         Left            =   1710
         TabIndex        =   26
         Top             =   5250
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlPartINLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   25
         Top             =   4680
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtTotalPriceIN 
         Height          =   465
         Left            =   7965
         TabIndex        =   27
         Top             =   5250
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   820
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   1005
         Left            =   150
         TabIndex        =   24
         Top             =   6360
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   1773
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
         Column(1)       =   "frmBarcodeProduction.frx":2F36
         Column(2)       =   "frmBarcodeProduction.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmBarcodeProduction.frx":30A2
         FormatStyle(2)  =   "frmBarcodeProduction.frx":31FE
         FormatStyle(3)  =   "frmBarcodeProduction.frx":32AE
         FormatStyle(4)  =   "frmBarcodeProduction.frx":3362
         FormatStyle(5)  =   "frmBarcodeProduction.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmBarcodeProduction.frx":34F2
      End
      Begin Threed.SSCommand cmdDeleteIN 
         Height          =   525
         Left            =   150
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":36CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblUnitIN 
         Height          =   375
         Left            =   3810
         TabIndex        =   32
         Top             =   5310
         Width           =   2565
      End
      Begin VB.Label lblTotalPriceIN 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         TabIndex        =   31
         Top             =   5310
         Width           =   1485
      End
      Begin VB.Label lblPartIN 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   4740
         Width           =   1485
      End
      Begin VB.Label lblQuantityIN 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   5310
         Width           =   1485
      End
      Begin Threed.SSCommand cmdSaveBarcodeIN 
         Height          =   525
         Left            =   9840
         TabIndex        =   28
         Top             =   5205
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSaveBarcode 
         Height          =   525
         Left            =   9840
         TabIndex        =   3
         Top             =   1950
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":3CFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2070
         Width           =   1485
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1500
         Width           =   1485
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6360
         TabIndex        =   20
         Top             =   2070
         Width           =   1485
      End
      Begin VB.Label lblUnit 
         Height          =   375
         Left            =   3810
         TabIndex        =   19
         Top             =   2070
         Width           =   2565
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9360
         TabIndex        =   18
         Top             =   900
         Width           =   735
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1740
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   6840
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":4332
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4050
         TabIndex        =   17
         Top             =   900
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":464C
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmBarcodeProduction.frx":4966
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   900
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmBarcodeProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_InventoryDoc As CInventoryDoc

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As INVENTORY_DOCTYPE

Private FileName As String
Private m_SumUnit As Double

Private m_Cd As Collection
Private DocAdd As Long
Private m_Parts As Collection
Private m_PartINs As Collection
Private m_Locations As Collection

Private m_LotItemOut As Collection
Private m_LotItemIn As Collection
'--------------------------------------------------
Private UnitID As Long
Private Multiple As Double
Private UnitName As String
Private UnitMName As String
'--------------------------------------------------
Private UnitIDIN As Long
Private MultipleIN As Double
Private UnitNameIN As String
Private UnitMNameIN As String
'------------------------------------------------------
Private CanSaveMatchUnit As Boolean
Private Sub CalculateSumPrice()
Dim Li As CLotItem
Dim Sum2 As Double
Dim TempUnitChangeName As String

   Sum2 = 0
   
   TempUnitChangeName = ""
   CanSaveMatchUnit = False
   
   For Each Li In m_LotItemOut
      If Li.Flag <> "D" Then
         Sum2 = Sum2 + Li.TX_AMOUNT
         
         If TempUnitChangeName = "" Then
            TempUnitChangeName = Li.UNIT_CHANGE_NAME
         End If
         If TempUnitChangeName <> Li.UNIT_CHANGE_NAME Then
            CanSaveMatchUnit = True
         End If
      End If
   Next Li
   
   For Each Li In m_LotItemIn
      If Li.Flag <> "D" Then
         Sum2 = Sum2 - Li.TX_AMOUNT
         
         If TempUnitChangeName = "" Then
            TempUnitChangeName = Li.UNIT_CHANGE_NAME
         End If
         If TempUnitChangeName <> Li.UNIT_CHANGE_NAME Then
            CanSaveMatchUnit = True
         End If
      End If
   Next Li
   
   txtTotalAmount.Text = Format(Sum2, "0.00")
   
End Sub
Private Sub MergeData()
Dim Li As CLotItem
Dim Sum2 As Double

   Sum2 = 0
   
   For Each Li In m_LotItemOut
      Call m_InventoryDoc.ImportExportItems.add(Li)
   Next Li
   
   For Each Li In m_LotItemIn
      Call m_InventoryDoc.ImportExportItems.add(Li)
   Next Li
   
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblDocumentDate, uctlDocumentDate, False) Then
      Exit Function
   End If
   
   If Not VerifyLockDate(uctlDocumentDate.ShowDate, m_InventoryDoc.GetFieldValue("DOCUMENT_DATE")) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
    If Not VerifyLockInventoryDate(uctlDocumentDate.ShowDate, m_InventoryDoc.GetFieldValue("DOCUMENT_DATE")) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Val(txtTotalAmount.Text) <> 0 And Not (CanSaveMatchUnit) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถบันทึกได้เนื่องจาก ---------------------> ยอดในการผลิตไม่เท่ากับยอดที่ผลิตได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not CheckUniqueNs(INVENTORY_DOC_NO, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_InventoryDoc.ShowMode = SHOW_ADD
   Call m_InventoryDoc.SetFieldValue("INVENTORY_DOC_ID", ID)
   If uctlDocumentDate.ShowDate <= 0 Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DATE", Now)
   Else
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DATE", uctlDocumentDate.ShowDate)
   End If
    
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_NO", txtDocumentNo.Text)
   If (DocumentType = ADJUST_DOCTYPE) Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", 1000) 'ผลิต
   Else
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   End If
   
   
   Call m_InventoryDoc.SetFieldValue("COMMIT_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("EXCEPTION_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("SALE_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("ADJUST_FLAG", "N")
   If DocumentType = IMPORT_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", "รับเข้าจาก BARCODE")
   ElseIf DocumentType = EXPORT_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", "เบิกออกจาก BARCODE")
   ElseIf DocumentType = ADJUST_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", "ผลิต")
   End If
   Call m_InventoryDoc.SetFieldValue("DEPARTMENT_ID", -1)
   Call m_InventoryDoc.SetFieldValue("CANCEL_FLAG", "N")
   If DocumentType = ADJUST_DOCTYPE Then
      Call m_InventoryDoc.SetFieldValue("BARCODE_JOB_FLAG", "Y")
   End If
   
   MergeData
      
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditInventoryDoc(m_InventoryDoc, IsOK, True, glbErrorLog) Then
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
Private Sub cmdAuto_Click()
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   If Len(txtDocumentNo.Text) > 0 Then
      SendKeys ("{TAB}")
      Exit Sub
   End If
   
   ID = 1000 'บังคับไปเลยว่าเป็น 1000 ซึ่งเป็นไปสั่งผลิต
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         txtDocumentNo.Text = Cd.GetFieldValue("PREFIX")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         txtDocumentNo.Text = txtDocumentNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
         Call m_InventoryDoc.SetFieldValue("RUNNING_NO", Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd)
         Call m_InventoryDoc.SetFieldValue("CONFIG_DOC_TYPE", ID)
         
         Call txtDocumentNo.SetSelectText(Len(txtDocumentNo.Text) - Cd.GetFieldValue("DIGIT_AMOUNT"), Cd.GetFieldValue("DIGIT_AMOUNT"))
      Else
         txtDocumentNo.Text = ""
      End If
   End If
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
   
   If Not ConfirmDelete(GridEX1.Value(4)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If ID1 <= 0 Then
         m_LotItemOut.Remove (ID2)
      Else
         m_LotItemOut.Item(ID2).Flag = "D"
      End If
      
      Call RefreshGrid(DocumentType, True)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdDeleteIN_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX2.Value(4)) Then
      Exit Sub
   End If
   
   ID2 = GridEX2.Value(2)
   ID1 = GridEX2.Value(1)
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If ID1 <= 0 Then
         m_LotItemIn.Remove (ID2)
      Else
         m_LotItemIn.Item(ID2).Flag = "D"
      End If
      
      Call RefreshGrid(DocumentType, True)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSaveBarcode_Click()
Dim IsOK As Boolean
Dim RealIndex As Long
Dim Mr As CMasterRef
Dim LocationName As String
   
   If Not VerifyCombo(lblPart, uctlPartLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblQuantity, txtQuantity, False) Then
      Exit Sub
   End If
   
   Dim EnpAddress As CLotItem
   
   Set EnpAddress = New CLotItem
   EnpAddress.Flag = "A"
   Call m_LotItemOut.add(EnpAddress)
   
   EnpAddress.PART_ITEM_ID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   For Each Mr In m_Locations
      If Mr.KEY_CODE = "01" Then
         EnpAddress.LOCATION_ID = Mr.KEY_ID
         EnpAddress.LOCATION_NAME = Mr.KEY_NAME
         Exit For
      End If
   Next Mr
   
   EnpAddress.TX_AMOUNT = Val(txtQuantity.Text) * Multiple
   EnpAddress.TOTAL_INCLUDE_PRICE = Val(txtTotalPrice.Text)
   EnpAddress.AVG_PRICE = MyDiffEx(EnpAddress.TOTAL_INCLUDE_PRICE, EnpAddress.TX_AMOUNT)
   
   EnpAddress.UNIT_TRAN_ID = UnitID
   EnpAddress.UNIT_MULTIPLE = Multiple
   EnpAddress.UNIT_TRAN_NAME = UnitName
   EnpAddress.UNIT_CHANGE_NAME = UnitMName
   
   EnpAddress.PART_NO = uctlPartLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartLookup.MyCombo.Text
   
   EnpAddress.TX_TYPE = "E"
   EnpAddress.MULTIPLIER = -1
   
   Set EnpAddress = Nothing
   
   Call RefreshGrid(DocumentType, False)
   
   txtQuantity.Text = ""
   txtTotalPrice.Text = ""
   uctlPartLookup.MyCombo.ListIndex = IDToListIndex(uctlPartLookup.MyCombo, -1)
   uctlPartLookup.SetFocus
End Sub
Private Sub cmdSaveBarcodeIN_Click()
Dim IsOK As Boolean
Dim RealIndex As Long
Dim Mr As CMasterRef
Dim LocationName As String
   
   If Not VerifyCombo(lblPartIN, uctlPartINLookup.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblQuantityIN, txtQuantityIN, False) Then
      Exit Sub
   End If
   
   Dim EnpAddress As CLotItem
   
   Set EnpAddress = New CLotItem
   EnpAddress.Flag = "A"
   Call m_LotItemIn.add(EnpAddress)
   
   EnpAddress.PART_ITEM_ID = uctlPartINLookup.MyCombo.ItemData(Minus2Zero(uctlPartINLookup.MyCombo.ListIndex))
   For Each Mr In m_Locations
      If Mr.KEY_CODE = "01" Then
         EnpAddress.LOCATION_ID = Mr.KEY_ID
         EnpAddress.LOCATION_NAME = Mr.KEY_NAME
         Exit For
      End If
   Next Mr
   
   EnpAddress.TX_AMOUNT = Val(txtQuantityIN.Text) * MultipleIN
   EnpAddress.TOTAL_INCLUDE_PRICE = Val(txtTotalPriceIN.Text)
   EnpAddress.AVG_PRICE = MyDiffEx(EnpAddress.TOTAL_INCLUDE_PRICE, EnpAddress.TX_AMOUNT)
   
   EnpAddress.UNIT_TRAN_ID = UnitIDIN
   EnpAddress.UNIT_MULTIPLE = MultipleIN
   EnpAddress.UNIT_TRAN_NAME = UnitNameIN
   EnpAddress.UNIT_CHANGE_NAME = UnitMNameIN
   
   EnpAddress.PART_NO = uctlPartINLookup.MyTextBox.Text
   EnpAddress.PART_DESC = uctlPartINLookup.MyCombo.Text
   
   EnpAddress.TX_TYPE = "I"
   EnpAddress.MULTIPLIER = 1
   
   Set EnpAddress = Nothing
   
   Call RefreshGrid(DocumentType, False)
   
   txtQuantityIN.Text = ""
   txtTotalPriceIN.Text = ""
   uctlPartINLookup.MyCombo.ListIndex = IDToListIndex(uctlPartINLookup.MyCombo, -1)
   uctlPartINLookup.SetFocus
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadStockBarcode(uctlPartLookup.MyCombo, m_Parts)
      Set uctlPartLookup.MyCollection = m_Parts
      
      Call LoadStockBarcode(uctlPartINLookup.MyCombo, m_PartINs)
      Set uctlPartINLookup.MyCollection = m_PartINs
            
      Call LoadMaster(Nothing, m_Locations, , , MASTER_LOCATION)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      m_InventoryDoc.ShowMode = SHOW_ADD
      
      uctlDocumentDate.ShowDate = Now
      Call cmdAuto_Click
      
      Call EnableForm(Me, True)
      m_HasModify = False
      
      uctlPartLookup.SetFocus
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
      'Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      'Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      'Call cmdEdit_Click
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
   
   Set m_InventoryDoc = Nothing
   Set m_Cd = Nothing
    Set m_Parts = Nothing
    Set m_PartINs = Nothing
   Set m_Locations = Nothing
   Set m_LotItemOut = Nothing
   Set m_LotItemIn = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1(Ind As Long)
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
   
   If Ind = IMPORT_DOCTYPE Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("สถานที่จัดเก็บ")
            
   ElseIf Ind = EXPORT_DOCTYPE Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("สถานที่จัดเก็บ")
            
   ElseIf Ind = TRANSFER_DOCTYPE Then
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("จากสถานที่จัดเก็บ")
      
      Set Col = GridEX1.Columns.add '9
      Col.Width = 1995
      Col.Caption = MapText("ไปสถานที่จัดเก็บ")
      
   ElseIf Ind = ADJUST_DOCTYPE Then
      
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1000
      Col.Caption = MapText("ประเภท")
      
      Set Col = GridEX1.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX1.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX1.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX1.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX1.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX1.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("สถานที่จัดเก็บ")
   End If
End Sub
Private Sub InitGrid2(Ind As Long)
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
   
   If Ind = IMPORT_DOCTYPE Then
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX2.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX2.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX2.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX2.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX2.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("สถานที่จัดเก็บ")
            
   ElseIf Ind = EXPORT_DOCTYPE Then
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX2.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX2.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX2.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX2.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX2.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("สถานที่จัดเก็บ")
            
   ElseIf Ind = TRANSFER_DOCTYPE Then
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX2.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX2.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX2.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX2.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX2.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("จากสถานที่จัดเก็บ")
      
      Set Col = GridEX2.Columns.add '9
      Col.Width = 1995
      Col.Caption = MapText("ไปสถานที่จัดเก็บ")
      
   ElseIf Ind = ADJUST_DOCTYPE Then
      
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1000
      Col.Caption = MapText("ประเภท")
      
      Set Col = GridEX2.Columns.add '3
      Col.Width = 1710
      Col.Caption = MapText("รหัสสต็อค")
   
      Set Col = GridEX2.Columns.add '4
      Col.Width = 4335
      Col.Caption = MapText("รายการ")
      
      Set Col = GridEX2.Columns.add '5
      Col.Width = 1785
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ปริมาณ")
   
      Set Col = GridEX2.Columns.add '6
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคา")
      
      Set Col = GridEX2.Columns.add '7
      Col.Width = 1620
      Col.TextAlignment = jgexAlignRight
      Col.Caption = MapText("ราคารวม")
      
      Set Col = GridEX2.Columns.add '8
      Col.Width = 1995
      Col.Caption = MapText("สถานที่จัดเก็บ")
   End If
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblTotalAmount, MapText("รวม"))
   Call InitNormalLabel(lblDocumentDate, MapText("วันที่เอกสาร"))
   Call InitNormalLabel(lblPart, MapText("สินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblQuantity, MapText("จำนวน"))
   Call InitNormalLabel(lblTotalPrice, MapText("ราคา"))
   
   Call InitNormalLabel(lblPartIN, MapText("สินค้า/วัตถุดิบ"))
   Call InitNormalLabel(lblQuantityIN, MapText("จำนวน"))
   Call InitNormalLabel(lblTotalPriceIN, MapText("ราคา"))
   
   Call InitNormalLabel(lblUnit, MapText(""))
   Call InitNormalLabel(lblUnitIN, MapText(""))
   
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtQuantity.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtQuantityIN.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   txtTotalAmount.Enabled = False
   
   txtTotalPrice.Enabled = False
   txtTotalPriceIN.Enabled = False
   
   uctlPartLookup.MyCombo.TabStop = False
   uctlPartINLookup.MyCombo.TabStop = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDeleteIN.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSaveBarcode.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSaveBarcodeIN.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdDelete, MapText("ลบ"))
   Call InitMainButton(cmdDeleteIN, MapText("ลบ"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdSaveBarcode, MapText("บันทึกเบิกไปผลิต"))
   Call InitMainButton(cmdSaveBarcodeIN, MapText("บันทึกรับจากผลิต"))
   
   
   Call InitGrid1(DocumentType)
   Call InitGrid2(DocumentType)
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip2.Font.Bold = True
   TabStrip2.Font.Name = GLB_FONT
   TabStrip2.Font.Size = 16
   
   Dim T As Object
   Dim TIN As Object
   TabStrip1.Tabs.Clear
   TabStrip2.Tabs.Clear
   If DocumentType = IMPORT_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = EXPORT_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = TRANSFER_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText(Doctype2Text(DocumentType))
      T.Tag = DocumentType & "-1"
   ElseIf DocumentType = ADJUST_DOCTYPE Then
      Set T = TabStrip1.Tabs.add()
      T.Caption = MapText("รายการเบิกไปผลิต")
      T.Tag = DocumentType & "-1"
      
      Set TIN = TabStrip2.Tabs.add()
      TIN.Caption = MapText("รายการรับจากผลิต")
      TIN.Tag = DocumentType & "-1"
   End If
End Sub

Private Function Doctype2Text(TempID As INVENTORY_DOCTYPE) As String
   If TempID = IMPORT_DOCTYPE Then
      Doctype2Text = "รายการนำเข้า"
   ElseIf TempID = EXPORT_DOCTYPE Then
      Doctype2Text = "รายการเบิกจ่าย"
   ElseIf TempID = TRANSFER_DOCTYPE Then
      Doctype2Text = "รายการโอนสต็อค"
   ElseIf TempID = ADJUST_DOCTYPE Then
      Doctype2Text = "รายการผลิต/ปรับยอด"
   End If
End Function

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
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Cd = New Collection
   Set m_Parts = New Collection
   Set m_PartINs = New Collection
   Set m_Locations = New Collection
   
   Set m_LotItemOut = New Collection
   Set m_LotItemIn = New Collection
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
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If m_LotItemOut Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If (DocumentType = ADJUST_DOCTYPE) Then
         Dim Aj As CLotItem
         If m_LotItemOut.Count <= 0 Then
            Exit Sub
         End If
         Set Aj = GetItem(m_LotItemOut, RowIndex, RealIndex)
         If Aj Is Nothing Then
            Exit Sub
         End If
         
         If Aj.TX_TYPE = "E" Then
            Values(1) = Aj.LOT_ITEM_ID
            Values(2) = RealIndex
            If Aj.TX_TYPE = "E" Then
               Values(3) = "ปรับลด"
            Else
               Values(3) = "ปรับเพิ่ม"
            End If
            Values(4) = Aj.PART_NO
            Values(5) = Aj.PART_DESC
            Values(6) = FormatNumber(MyDiff(Aj.TX_AMOUNT, Aj.UNIT_MULTIPLE))
            Values(7) = FormatNumber(Aj.AVG_PRICE * Aj.UNIT_MULTIPLE)
            Values(8) = FormatNumber(Aj.TOTAL_INCLUDE_PRICE)
            Values(9) = Aj.LOCATION_NAME
         End If
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub RefreshGrid(Ind As INVENTORY_DOCTYPE, Flag As Boolean)

   GridEX1.ItemCount = CountItem(m_LotItemOut)
   GridEX1.Rebind
   
   GridEX2.ItemCount = CountItem(m_LotItemIn)
   GridEX2.Rebind

   Call CalculateSumPrice
   If Flag Then
      m_HasModify = Flag
   End If
End Sub
Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      If m_LotItemOut Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      If (DocumentType = ADJUST_DOCTYPE) Then
         Dim Aj As CLotItem
         If m_LotItemIn.Count <= 0 Then
            Exit Sub
         End If
         Set Aj = GetItem(m_LotItemIn, RowIndex, RealIndex)
         If Aj Is Nothing Then
            Exit Sub
         End If
         
         If Aj.TX_TYPE = "I" Then
            Values(1) = Aj.LOT_ITEM_ID
            Values(2) = RealIndex
            If Aj.TX_TYPE = "E" Then
               Values(3) = "ปรับลด"
            Else
               Values(3) = "ปรับเพิ่ม"
            End If
            Values(4) = Aj.PART_NO
            Values(5) = Aj.PART_DESC
            Values(6) = FormatNumber(MyDiff(Aj.TX_AMOUNT, Aj.UNIT_MULTIPLE))
            Values(7) = FormatNumber(Aj.AVG_PRICE * Aj.UNIT_MULTIPLE)
            Values(8) = FormatNumber(Aj.TOTAL_INCLUDE_PRICE)
            Values(9) = Aj.LOCATION_NAME
         End If
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.Tag = DocumentType & "-1" Then
      Call InitGrid1(DocumentType)
      Call RefreshGrid(DocumentType, False)
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtDocumentNo_LostFocus()
   If Not CheckUniqueNs(INVENTORY_DOC_NO, txtDocumentNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
End Sub
Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX2.Width = ScaleWidth - 2 * GridEX2.Left
   TabStrip1.Width = GridEX1.Width
   TabStrip2.Width = GridEX2.Width
   GridEX2.Height = ScaleHeight - GridEX2.Top - 620
   cmdPrint.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
   cmdPrint.Left = cmdOK.Left - cmdPrint.Width - 50
End Sub
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

Private Sub uctlPartINLookup_Change()
Dim PartItemID As Long
Dim Pi As CStockCode
   
   PartItemID = uctlPartINLookup.MyCombo.ItemData(Minus2Zero(uctlPartINLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(PartItemID)))
      UnitIDIN = Pi.UNIT_ID
      MultipleIN = Pi.UNIT_AMOUNT
      UnitNameIN = Pi.UNIT_NAME
      UnitMNameIN = Pi.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnitIN, UnitNameIN & " X " & MultipleIN & " " & UnitMNameIN)
   End If
   m_HasModify = True

End Sub

Private Sub uctlPartLookup_Change()
Dim PartItemID As Long
Dim Pi As CStockCode
   
   PartItemID = uctlPartLookup.MyCombo.ItemData(Minus2Zero(uctlPartLookup.MyCombo.ListIndex))
   If PartItemID > 0 Then
      Set Pi = GetObject("CStockCode", m_Parts, Trim(Str(PartItemID)))
      UnitID = Pi.UNIT_ID
      Multiple = Pi.UNIT_AMOUNT
      UnitName = Pi.UNIT_NAME
      UnitMName = Pi.UNIT_CHANGE_NAME
      
      Call InitNormalLabel(lblUnit, UnitName & " X " & Multiple & " " & UnitMName)
   End If
   m_HasModify = True
End Sub
