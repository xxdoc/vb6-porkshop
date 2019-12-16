VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddReturnItem 
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13785
   Icon            =   "frmAddReturnItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   13785
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10440
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   18415
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   7815
         Left            =   150
         TabIndex        =   3
         Top             =   1590
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   13785
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
         Column(1)       =   "frmAddReturnItem.frx":27A2
         Column(2)       =   "frmAddReturnItem.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddReturnItem.frx":290E
         FormatStyle(2)  =   "frmAddReturnItem.frx":2A6A
         FormatStyle(3)  =   "frmAddReturnItem.frx":2B1A
         FormatStyle(4)  =   "frmAddReturnItem.frx":2BCE
         FormatStyle(5)  =   "frmAddReturnItem.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddReturnItem.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   7815
         Left            =   6360
         TabIndex        =   5
         Top             =   1590
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   13785
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
         Column(1)       =   "frmAddReturnItem.frx":2F36
         Column(2)       =   "frmAddReturnItem.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddReturnItem.frx":30A2
         FormatStyle(2)  =   "frmAddReturnItem.frx":31FE
         FormatStyle(3)  =   "frmAddReturnItem.frx":32AE
         FormatStyle(4)  =   "frmAddReturnItem.frx":3362
         FormatStyle(5)  =   "frmAddReturnItem.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddReturnItem.frx":34F2
      End
      Begin PorkShop.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   7260
         TabIndex        =   1
         Top             =   930
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblF7 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         Top             =   4800
         Width           =   795
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   5600
         TabIndex        =   4
         Top             =   5190
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReturnItem.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   12060
         TabIndex        =   2
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReturnItem.frx":39E4
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   11
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label lblDocumentDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5970
         TabIndex        =   10
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5400
         TabIndex        =   6
         Top             =   9660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddReturnItem.frx":3CFE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7050
         TabIndex        =   7
         Top             =   9660
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddReturnItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_PO As CBillingDoc

Public CusID As Long
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection
Public DocumentType As SELL_BILLING_DOCTYPE

Private FileName As String
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection

Public PO_ID As Long
Public PO_NO As String


Private Sub PopulateDestColl()
Dim Ri As CDocItem
Dim D As CDocItem

   If TempCollection Is Nothing Then
      Exit Sub
   End If
   
   For Each Ri In TempCollection
      Set D = New CDocItem
      
      If Ri.Flag <> "D" Then
         Call D.CopyObject(1, Ri)
         Call m_TempCol2.add(D)
      End If
      
      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, TempID As Long) As Boolean
Dim D As CDocItem
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.GetFieldValue("PO_ID") = TempID Then
         Found = True
      End If
   Next D
   
   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim BD As CBillingDoc

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   
   While Not Rs.EOF
      Set BD = New CBillingDoc
      Call BD.PopulateFromRS(1002, Rs)
      
      If Not IsIn(m_TempCol2, BD.BILLING_DOC_ID) Then
         Call TempCol.add(BD)
      End If

      Set BD = Nothing
      Rs.MoveNext
   Wend
   
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)

      m_PO.BILLING_DOC_ID = -1
      m_PO.FROM_DATE = uctlFromDate.ShowDate
      m_PO.TO_DATE = uctlToDate.ShowDate
      m_PO.APAR_MAS_ID = CusID
      If DocumentType = RETURN_DOCTYPE Then
         m_PO.DOCUMENT_TYPE = INVOICE_DOCTYPE
      ElseIf DocumentType = S_RETURN_DOCTYPE Then
         m_PO.DOCUMENT_TYPE = S_INVOICE_DOCTYPE
      End If
      m_PO.ORDER_TYPE = 2
      Call m_PO.QueryData(1002, m_Rs, ItemCount)
   End If
   
   If m_Rs.RecordCount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.ItemCount = 0
      GridEX1.Rebind
   End If
   
   If Not (m_TempCol2 Is Nothing) Then
      GridEX2.ItemCount = m_TempCol2.Count
      GridEX2.Rebind
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call PopulateTempColl

   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CBillingDoc
Dim OKClick As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim Poi As CDocItem
Dim Di As CDocItem
Dim IsOK As Boolean
Dim TempColl As Collection
Dim GetUpdate As CDocItem
Set TempColl = New Collection
   
   Set TempRs = New ADODB.Recordset
   
   If ID > 0 Then
      Set L = TempCol1(ID)
      L.DOCUMENT_TYPE = INVOICE_DOCTYPE
      L.QueryFlag = 1
      
      Call GetDoitem(L)
      
      For Each Poi In L.DocItems
         Set Di = New CDocItem
         Call Di.CopyObject(1, Poi)
         
         Di.Flag = "A"
         
         Call Di.SetFieldValue("PO_ID", Poi.GetFieldValue("BILLING_DOC_ID"))
         Call Di.SetFieldValue("PO_NO", L.DOCUMENT_NO)
         
         PO_ID = Poi.GetFieldValue("BILLING_DOC_ID")
         PO_NO = L.DOCUMENT_NO
         
         If ShowMode = SHOW_VIEW_ONLY Then
            Exit Sub
         End If
         
         Call TempCol2.add(Di)
         
         Set Di = Nothing
      Next Poi
      Call TempCol1.Remove(ID)
   End If
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long

   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
      TempCol1(j).IncludeFlag = True
      Call TempCol2.add(TempCol1(j))
   Next j
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long

   m_HasModify = True
      
   If GridEX2.ItemCount <= 0 Then
      
      TempID = GridEX1.Row
      Call CopyItem(m_TempCol1, m_TempCol2, TempID)
   
      GridEX1.ItemCount = m_TempCol1.Count
      GridEX1.Rebind
      
      GridEX2.ItemCount = m_TempCol2.Count
      GridEX2.Rebind
      
      Call cmdOK_Click
      
   End If
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
Dim D As CDocItem
Dim Ri As CDocItem


   For Each D In m_TempCol2
      Set Ri = New CDocItem

      If (D.Flag = "A") Then
         Call Ri.CopyObject(1, D)
         Call Ri.SetFieldValue("PO_ID", D.GetFieldValue("PO_ID"))
         Call Ri.SetFieldValue("PO_NO", D.GetFieldValue("PO_NO"))
         Ri.Flag = "A"
         Call TempCollection.add(Ri)
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
         m_PO.QueryFlag = 0
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_PO.QueryFlag = 0
         uctlFromDate.ShowDate = DateAdd("YYYY", -1, Now)
         uctlToDate.ShowDate = Now
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
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdSelect_Click
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
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX2.Width = ScaleWidth - GridEX1.Left - GridEX1.Width - 1200
   GridEX2.Height = ScaleHeight - GridEX2.Top - cmdOK.Height - 100
   GridEX1.Height = GridEX2.Height
   cmdOK.Top = ScaleHeight - cmdOK.Height - 50
   cmdExit.Top = ScaleHeight - cmdExit.Height - 50
   cmdExit.Left = ScaleWidth / 2 + 50
   cmdOK.Left = ScaleWidth / 2 - cmdOK.Width - 50
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_PO = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdSelect_Click
End Sub

Private Sub GridEX2_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''debug.print ColIndex & " " & NewColWidth
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
   Col.Width = 1575
   Col.Caption = MapText("วันที่เอกสาร")

   Set Col = GridEX1.Columns.add '4
   Col.Width = 1740
   Col.Caption = MapText("หมายเลขเอกสาร")

   Set Col = GridEX1.Columns.add '5
   Col.Width = 1740
   Col.Caption = MapText("วันที่ครบกำหนด")
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
   Col.Width = 2500
   Col.Caption = MapText("รายละเอียด")
   
   Set Col = GridEX2.Columns.add '4
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน")
   
   Set Col = GridEX2.Columns.add '5
   Col.Width = 1500
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคารวม")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblDocumentDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblF7, MapText("F7"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา"))
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
   
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   
   Set m_PO = New CBillingDoc
   
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
   Values(5) = DateToStringExtEx2(CR.Due_Date)
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
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

   Dim CR As CDocItem
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.GetFieldValue("DOC_ITEM_ID")
   Values(2) = RealIndex
   Values(3) = CR.ShowDescText
   Values(4) = FormatNumber(CR.GetFieldValue("ITEM_AMOUNT"))
   Values(5) = FormatNumber(CR.GetFieldValue("TOTAL_PRICE"))
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetUpDateFromPo(ID As Long, TmpCol As Collection)
Dim Doc As CDocItem
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long
Set Doc = New CDocItem
Set m_Rs = New ADODB.Recordset
Set TmpCol = New Collection

   Call Doc.SetFieldValue("PO_ID", ID)
   Call Doc.QueryData(2, m_Rs, ItemCount)
   
   While Not m_Rs.EOF
      Set Doc = New CDocItem
      Call Doc.PopulateFromRS(2, m_Rs)
      Call TmpCol.add(Doc, Trim(Str(Doc.GetFieldValue("PART_ITEM_ID"))))
      m_Rs.MoveNext
   Wend
   
   Set m_Rs = Nothing
   Set Doc = Nothing
End Sub
Public Sub GetDoitem(Ua As CBillingDoc)
Dim Gr As CDocItem
Dim m_Rs1 As ADODB.Recordset
Dim iCount As Long

      Set m_Rs1 = New ADODB.Recordset
      Set Gr = New CDocItem
      'Call Gr.SetFieldValue("DOC_ITEM_TYPE", Ua.GetFieldValue("DOC_ITEM_TYPE"))          'ไม่แน่ใจว่าไม่มีได้หรือเปล่า
      Call Gr.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.DocItems = Nothing
      Set Ua.DocItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CDocItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         
         Gr.Flag = "I"
         Call Ua.DocItems.add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      Set Gr = Nothing
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
Private Sub GridEX2_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

