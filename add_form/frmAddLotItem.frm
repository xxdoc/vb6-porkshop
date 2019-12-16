VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAddLotItem 
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   Icon            =   "frmAddLotItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11985
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   9000
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15875
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSFrame SSFrame2 
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   5640
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Rebate.uctlTextBox txtAmount 
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   873
         End
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3615
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "กด Space Bar หรือ Double Click เพื่อเลือกเอกสาร "
         Top             =   1920
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   6376
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
         Column(1)       =   "frmAddLotItem.frx":27A2
         Column(2)       =   "frmAddLotItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddLotItem.frx":290E
         FormatStyle(2)  =   "frmAddLotItem.frx":2A6A
         FormatStyle(3)  =   "frmAddLotItem.frx":2B1A
         FormatStyle(4)  =   "frmAddLotItem.frx":2BCE
         FormatStyle(5)  =   "frmAddLotItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddLotItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1140
         TabIndex        =   0
         Top             =   1410
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Rebate.uctlDate uctlDocumentDate 
         Height          =   405
         Left            =   6180
         TabIndex        =   1
         Top             =   1410
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   1815
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "กด Space Bar หรือ Double Click เพื่อเลือกเอกสาร "
         Top             =   6480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3201
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
         Column(1)       =   "frmAddLotItem.frx":2F36
         Column(2)       =   "frmAddLotItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddLotItem.frx":30A2
         FormatStyle(2)  =   "frmAddLotItem.frx":31FE
         FormatStyle(3)  =   "frmAddLotItem.frx":32AE
         FormatStyle(4)  =   "frmAddLotItem.frx":3362
         FormatStyle(5)  =   "frmAddLotItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddLotItem.frx":34F2
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   735
         Left            =   6840
         TabIndex        =   13
         Top             =   5640
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1296
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Rebate.uctlTextBox txtAmount2 
            Height          =   495
            Left            =   2160
            TabIndex        =   14
            Top             =   120
            Width           =   2655
            _ExtentX        =   5741
            _ExtentY        =   873
         End
         Begin Rebate.uctlTextBox txtImportLotItemID 
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   2055
            _ExtentX        =   4471
            _ExtentY        =   873
         End
      End
      Begin Rebate.uctlTextBox txtPartItemID 
         Height          =   495
         Left            =   1140
         TabIndex        =   16
         Top             =   840
         Width           =   1815
         _ExtentX        =   5741
         _ExtentY        =   873
      End
      Begin Rebate.uctlTextBox txtPartNo 
         Height          =   495
         Left            =   6180
         TabIndex        =   18
         Top             =   840
         Width           =   1815
         _ExtentX        =   5741
         _ExtentY        =   873
      End
      Begin VB.Label lblPartNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   19
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label lblPartItemID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   915
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10140
         TabIndex        =   2
         Top             =   930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddLotItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5010
         TabIndex        =   8
         Top             =   1440
         Width           =   1035
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4200
         TabIndex        =   4
         Top             =   8340
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddLotItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5850
         TabIndex        =   5
         Top             =   8340
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddLotItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public TempCollection As Collection

Private m_LotItem As CLotItem

Public DocumentDate As Date
Public CanEditPartFlag As Boolean

Public TempLocationID As Long
Public TempPartItemID As Long
Public TempTxAmount  As Double

Public OneToOne As Boolean
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      Set m_LotItem = New CLotItem
      m_LotItem.LOT_ITEM_ID = -1
      m_LotItem.PART_ITEM_ID = Val(txtPartItemID.Text)
      m_LotItem.PART_NO = txtPartNo.Text
      m_LotItem.LOCATION_ID = TempLocationID
      m_LotItem.FROM_DOC_DATE = uctlFromDate.ShowDate
      m_LotItem.TO_DOC_DATE = uctlDocumentDate.ShowDate
      m_LotItem.COUNT_AMOUNT = "Y"
      Call m_LotItem.QueryData(6, m_Rs, ItemCount, True)
   End If
      
   Call InitGrid1
   Call InitGrid2
   
   If ItemCount > 0 Or m_Rs.RecordCount > 0 Then
      GridEX1.ItemCount = m_Rs.RecordCount
      GridEX1.Rebind
   End If
   
   If TempCollection.Count > 0 Then
      GridEX2.ItemCount = TempCollection.Count
      GridEX2.Rebind
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call InitGrid1
      
      uctlFromDate.ShowDate = DateAdd("M", -1, DocumentDate)
      
      uctlDocumentDate.ShowDate = DocumentDate
      Call QueryData(True)
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 32 Then
      Call GridEX1_DblClick
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdSelect_Click
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
   
   Set m_LotItem = Nothing

End Sub

Private Sub GridEX1_DblClick()
Dim ID  As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If OneToOne And TempCollection.Count > 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่สามารถเพิ่มรายการได้เนื่องจากระบบตั้งใว้สำหรับการโอนแบบ 1 ต่อ 1 เท่านั้น"
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   ID = GridEX1.Value(1)
   If ID > 0 Then
      txtAmount.Text = GridEX1.Value(9)
      SSFrame2.Visible = True
      txtAmount.SetFocus
   End If
   
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
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
   Col.Caption = MapText("รหัสเอกสาร")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("หมายเลข")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 0
   Col.Caption = MapText("ID วัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '6
   Col.Width = 1500
   Col.Caption = MapText("รหัสวัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '7
   Col.Width = ScaleWidth - 10500
   Col.Caption = MapText("วัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '8
   Col.Width = 1500
   Col.Caption = MapText("จำนวนจริง")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '9
   Col.Width = 1500
   Col.Caption = MapText("คงเหลือ")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.add '10
   Col.Width = 0
   Col.Caption = MapText("รหัส คลัง")
   
   Set Col = GridEX1.Columns.add '11
   Col.Width = 0
   Col.Caption = MapText("รหัส หน่วย")
   
   Set Col = GridEX1.Columns.add '12
   Col.Width = 0
   Col.Caption = MapText("รหัสประเภทเอกสาร")
   
   Set Col = GridEX1.Columns.add '13
   Col.Width = 1500
   Col.Caption = MapText("ราคาต่อหน่วย")
   Col.TextAlignment = jgexAlignRight
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
   Col.Caption = MapText("รหัสเอกสาร")
   
   Set Col = GridEX2.Columns.add '3
   Col.Width = 1500
   Col.Caption = MapText("วันที่เอกสาร")
   
   Set Col = GridEX2.Columns.add '4
   Col.Width = 2000
   Col.Caption = MapText("หมายเลข")
   
   Set Col = GridEX2.Columns.add '5
   Col.Width = 0
   Col.Caption = MapText("ID วัตถุดิบ")
   
   Set Col = GridEX2.Columns.add '6
   Col.Width = 1500
   Col.Caption = MapText("รหัสวัตถุดิบ")
   
   Set Col = GridEX2.Columns.add '7
   Col.Width = ScaleWidth - 7500
   Col.Caption = MapText("วัตถุดิบ")
   
   Set Col = GridEX2.Columns.add '8
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX2.Columns.add '9
   Col.Width = 0
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("รหัสหน่วย")
   
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame3.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblPartItemID, MapText("ID"))
   Call InitNormalLabel(lblPartNo, MapText("รหัสวัตถุดิบ"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   txtImportLotItemID.Enabled = False
   
   txtPartItemID.Text = TempPartItemID
   
   If Not (CanEditPartFlag) Then
      txtPartItemID.Enabled = False
      txtPartNo.Enabled = False
   End If
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   
   Call txtPartNo.SetKeySearch("STOCK_NO")
   
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
   MasterInd = "6"
   Set m_LotItem = New CLotItem
   MasterInd = "1"
   
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Height = ScaleHeight - 2500 - GridEX2.Height - 200 - SSFrame2.Height - 100
   GridEX1.Width = ScaleWidth - 500
   GridEX2.Width = ScaleWidth - 500
   SSFrame2.Top = GridEX1.Top + GridEX1.Height + 100
   SSFrame3.Top = SSFrame2.Top
   SSFrame3.Left = ScaleWidth - SSFrame3.Width = 250
   GridEX2.Top = SSFrame2.Top + SSFrame2.Height + 200
   cmdOK.Top = ScaleHeight - cmdOK.Height - 50
   cmdExit.Top = ScaleHeight - cmdExit.Height - 50
   cmdExit.Left = ScaleWidth / 2 + 50
   cmdOK.Left = ScaleWidth / 2 - cmdOK.Width - 50
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I  As Byte
Dim Lk As CLotItemLink
Dim TempNo As String
Dim TempPartNo  As String
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If
   
   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_LotItem.PopulateFromRS(6, m_Rs)
   
   I = 0
   I = I + 1
   Values(I) = m_LotItem.LOT_ITEM_ID
   I = I + 1
   Values(I) = m_LotItem.INVENTORY_DOC_ID
   I = I + 1
   Values(I) = DateToStringExtEx2(m_LotItem.DOCUMENT_DATE)
   I = I + 1
   Values(I) = m_LotItem.DOCUMENT_NO
   TempNo = m_LotItem.DOCUMENT_NO
   I = I + 1
   Values(I) = m_LotItem.PART_ITEM_ID
   TempPartNo = m_LotItem.PART_ITEM_ID
   I = I + 1
   Values(I) = m_LotItem.PART_NO
   I = I + 1
   Values(I) = m_LotItem.PART_DESC
   I = I + 1
   Values(I) = m_LotItem.LOT_ITEM_AMOUNT
   I = I + 1
   Values(I) = m_LotItem.LOT_ITEM_AMOUNT - m_LotItem.TX_AMOUNT
   I = I + 1
   Values(I) = m_LotItem.LOCATION_ID
   I = I + 1
   Values(I) = m_LotItem.UNIT_ID
   I = I + 1
   Values(I) = m_LotItem.DOCUMENT_TYPE
   I = I + 1
   Values(I) = m_LotItem.AVG_PRICE
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX2_DblClick()
Dim ID  As Long
   
   If Not VerifyGrid(GridEX2.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX2.Value(1)
   If ID > 0 Then
      txtImportLotItemID.Text = Val(ID)
      txtAmount2.Text = GridEX2.Value(8)
      SSFrame3.Visible = True
      txtAmount2.SetFocus
   End If
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim I  As Byte
Dim Lk As CLotItemLink

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TempCollection Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   If TempCollection.Count <= 0 Then
      Exit Sub
   End If
   Set Lk = GetItem(TempCollection, RowIndex, RealIndex)
   If Lk Is Nothing Then
      Exit Sub
   End If
   
   I = 0
   I = I + 1
   Values(I) = Lk.IMPORT_LOT_ITEM_ID
   I = I + 1
   Values(I) = m_LotItem.INVENTORY_DOC_ID
   I = I + 1
   Values(I) = DateToStringExtEx2(Lk.DOCUMENT_DATE)
   I = I + 1
   Values(I) = Lk.DOCUMENT_NO
   I = I + 1
   Values(I) = Lk.PART_ITEM_ID
   I = I + 1
   Values(I) = Lk.PART_NO
   I = I + 1
   Values(I) = Lk.PART_DESC
   I = I + 1
   Values(I) = Lk.IMPORT_AMOUNT
   I = I + 1
   Values(I) = Lk.UNIT_ID
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
 Private Sub txtAmount_LostFocus()
Dim Lk As CLotItemLink
Dim TempID As Long
   
   Set Lk = GetObject("CLotItemLink", TempCollection, Trim(GridEX1.Value(1) & "-" & GridEX1.Value(5)), False)
   If Lk Is Nothing Then
      Set Lk = New CLotItemLink
      Lk.Flag = "A"
      Lk.IMPORT_LOT_ITEM_ID = GridEX1.Value(1)
      Lk.DOCUMENT_DATE = GridEX1.Value(3)
      Lk.DOCUMENT_NO = GridEX1.Value(4)
      Lk.PART_ITEM_ID = GridEX1.Value(5)
      Lk.PART_NO = GridEX1.Value(6)
      Lk.PART_DESC = GridEX1.Value(7)
      Lk.LOT_ITEM_AMOUNT = Val(GridEX1.Value(8))
      Lk.IMPORT_AMOUNT = Val(txtAmount.Text)
      Lk.UNIT_ID = GridEX1.Value(11)
      Lk.AVG_PRICE = GridEX1.Value(13)
      
      Lk.MAIN_IMPORT_LOT_ITEM_ID = Lk.IMPORT_LOT_ITEM_ID
      
      TempID = Lk.IMPORT_LOT_ITEM_ID
      If GridEX1.Value(12) = 1000 Then          'ถ้าประเภทเอกสาร เป็น ใบสั่งผลิตแล้ว
         Call glbDaily.GetNextLotItemID(TempID, GridEX1.Value(2), -1)
      Else
         Call glbDaily.GetNextLotItemID(TempID, GridEX1.Value(2), GridEX1.Value(5))
      End If
      If TempID > 0 Then
         Lk.MAIN_IMPORT_LOT_ITEM_ID = TempID
      End If
      
      Call TempCollection.add(Lk, Trim(GridEX1.Value(1) & "-" & GridEX1.Value(5)))
      
      GridEX2.ItemCount = TempCollection.Count
      GridEX2.Rebind
      
      SSFrame2.Visible = False
      GridEX1.SetFocus
   Else
      glbErrorLog.LocalErrorMsg = "มีข้อมูลที่เลือกแล้ว"
      glbErrorLog.ShowUserError
      txtAmount.Text = ""
      SSFrame2.Visible = False
      GridEX1.SetFocus
   End If
End Sub
Private Sub GridEX2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
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
Private Sub txtAmount2_LostFocus()
Dim Lk As CLotItemLink
   For Each Lk In TempCollection
      If Lk.IMPORT_LOT_ITEM_ID = Val(txtImportLotItemID.Text) Then
         Lk.Flag = "E"
         Lk.IMPORT_AMOUNT = Val(txtAmount2.Text)
      End If
   Next Lk
   SSFrame3.Visible = False
   GridEX2.ItemCount = TempCollection.Count
   GridEX2.Rebind
End Sub
