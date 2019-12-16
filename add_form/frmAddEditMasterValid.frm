VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditMasterValid 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditMasterValid.frx":0000
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
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   4
         Top             =   2040
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
      Begin PorkShop.uctlTextBox txtMasterValidNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   840
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
         Height          =   4965
         Left            =   150
         TabIndex        =   5
         Top             =   2600
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8758
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
         Column(1)       =   "frmAddEditMasterValid.frx":27A2
         Column(2)       =   "frmAddEditMasterValid.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditMasterValid.frx":290E
         FormatStyle(2)  =   "frmAddEditMasterValid.frx":2A6A
         FormatStyle(3)  =   "frmAddEditMasterValid.frx":2B1A
         FormatStyle(4)  =   "frmAddEditMasterValid.frx":2BCE
         FormatStyle(5)  =   "frmAddEditMasterValid.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditMasterValid.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtMasterValidDesc 
         Height          =   435
         Left            =   5760
         TabIndex        =   1
         Top             =   840
         Width           =   6020
         _ExtentX        =   10610
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   7920
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblMasterValidDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   15
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   570
         TabIndex        =   14
         Top             =   1380
         Width           =   1155
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterValid.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   10
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterValid.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterValid.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblMasterValidNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditMasterValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_MasterValid As CMasterValid

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As MASTER_REBATE_AREA
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
   
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      
      m_MasterValid.MASTER_VALID_ID = ID
      If Not glbDaily.QueryMasterValid(m_MasterValid, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_MasterValid.PopulateFromRS(1, m_Rs)
      
      txtMasterValidNo.Text = m_MasterValid.MASTER_VALID_NO
      txtMasterValidDesc.Text = m_MasterValid.MASTER_VALID_DESC
      uctlFromDate.ShowDate = m_MasterValid.VALID_FROM
      uctlToDate.ShowDate = m_MasterValid.VALID_TO
      
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

   If Not VerifyTextControl(lblMasterValidNo, txtMasterValidNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If

   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   If Not CheckUniqueNs(MASTER_VALID_NO, txtMasterValidNo.Text, ID, Trim(Str(DocumentType))) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtMasterValidNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call txtMasterValidNo.SetFocus
      Exit Function
   End If
   
   If Not CheckUniqueCollection() Then
      Call GridEX1.SetFocus
      Exit Function
   End If
   
   m_MasterValid.AddEditMode = ShowMode
   m_MasterValid.MASTER_VALID_ID = ID
   m_MasterValid.MASTER_VALID_NO = txtMasterValidNo.Text
   m_MasterValid.MASTER_VALID_DESC = txtMasterValidDesc.Text
   m_MasterValid.VALID_FROM = uctlFromDate.ShowDate
   m_MasterValid.VALID_TO = uctlToDate.ShowDate
   m_MasterValid.MASTER_VALID_TYPE = DocumentType
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditMasterValid(m_MasterValid, IsOK, True, glbErrorLog) Then
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
Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditRebateCondition.ParentForm = Me
      Set frmAddEditRebateCondition.TempCollection = m_MasterValid.CollRebateCondition
      frmAddEditRebateCondition.ParentIndexTab = TabStrip1.SelectedItem.Index
      frmAddEditRebateCondition.ShowMode = SHOW_ADD
      frmAddEditRebateCondition.HeaderText = "เพิ่ม" & RebateType2Text(DocumentType)
      Load frmAddEditRebateCondition
      frmAddEditRebateCondition.Show 1

      OKClick = frmAddEditRebateCondition.OKClick

      Unload frmAddEditRebateCondition
      Set frmAddEditRebateCondition = Nothing
      
   End If
   
   If OKClick Then
      Call RefreshGrid
   End If

   If OKClick Then
      m_HasModify = True
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
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_MasterValid.CollRebateCondition.Remove (ID2)
      Else
         m_MasterValid.CollRebateCondition.Item(ID2).Flag = "D"
      End If
      m_HasModify = True
   End If
   
   Call RefreshGrid
   
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditRebateCondition.ParentForm = Me
      Set frmAddEditRebateCondition.TempCollection = m_MasterValid.CollRebateCondition
      frmAddEditRebateCondition.ID = ID
      frmAddEditRebateCondition.ParentIndexTab = TabStrip1.SelectedItem.Index
      frmAddEditRebateCondition.ShowMode = SHOW_EDIT
      frmAddEditRebateCondition.HeaderText = "แก้ไข" & RebateType2Text(DocumentType)
      Load frmAddEditRebateCondition
      frmAddEditRebateCondition.Show 1

      OKClick = frmAddEditRebateCondition.OKClick

      Unload frmAddEditRebateCondition
      Set frmAddEditRebateCondition = Nothing
   End If
      
   If OKClick Then
      Call RefreshGrid
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
      ID = m_MasterValid.MASTER_VALID_ID
      Set m_MasterValid = New CMasterValid
      m_MasterValid.QueryFlag = 1
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
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_MasterValid.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_MasterValid.QueryFlag = 0
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
   
   Set m_MasterValid = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn
Dim I As Byte
   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   I = 4
   
   Set Col = GridEX1.Columns.add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = (ScaleWidth - 600) / I
   Col.Caption = MapText("ประเภทลูกค้า")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = (ScaleWidth - 600) / I
   Col.Caption = MapText("กลุ่มสินค้า")

   Set Col = GridEX1.Columns.add '5
   Col.Width = (ScaleWidth - 600) / I
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("LEVEL (0 = ตัวเอง)")
   
   Set Col = GridEX1.Columns.add '4
   Col.Width = (ScaleWidth - 600) / I
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("Percent")
End Sub
Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblMasterValidNo, MapText("หมายเลข"))
   Call InitNormalLabel(lblMasterValidDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
   Call txtMasterValidNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("เงื่อนไข REBATE")
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
   Set m_MasterValid = New CMasterValid
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   If TabStrip1.SelectedItem.Index = 5 Then
'      RowBuffer.RowStyle = RowBuffer.Value(7)
'   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim CR As CRebateCondition

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If GetCollection(TabStrip1.SelectedItem.Index) Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      If GetCollection(TabStrip1.SelectedItem.Index).Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(GetCollection(TabStrip1.SelectedItem.Index), RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.REBATE_CONDITION_ID
      Values(2) = RealIndex
      Values(3) = CR.CUSTOMER_TYPE_NAME
      Values(4) = CR.STOCK_GROUP_NAME
      Values(5) = CR.REBATE_LEVEL
      Values(6) = FormatNumberToNull(CR.REBATE_PERCENT, 1)
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
   
  If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Index))
      GridEX1.Rebind
   End If
End Sub

Private Sub txtMasterValidDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtMasterValidNo_Change()
   m_HasModify = True
End Sub
Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
Public Sub RefreshGrid(Optional Index As Long = -1)
   If Index > 0 Then
      GridEX1.ItemCount = CountItem(GetCollection(Index))
      GridEX1.Rebind
      m_HasModify = True
   Else
      GridEX1.ItemCount = CountItem(GetCollection(TabStrip1.SelectedItem.Index))
      GridEX1.Rebind
   End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Top = TabStrip1.Top + TabStrip1.Height
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
Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Function GetCollection(Index As Long) As Collection
   If Index = 1 Then
      Set GetCollection = m_MasterValid.CollRebateCondition
   End If
End Function
Private Function CheckUniqueCollection() As Boolean
On Error GoTo ErrorHandler
Dim TempRbCdt As CRebateCondition
Dim RbCdt As CRebateCondition
Dim CollTemp As Collection
Set CollTemp = New Collection
   For Each TempRbCdt In m_MasterValid.CollRebateCondition
      If TempRbCdt.Flag <> "D" Then
         Call CollTemp.add(TempRbCdt, TempRbCdt.CUSTOMER_TYPE & "-" & TempRbCdt.STOCK_GROUP & "-" & TempRbCdt.REBATE_LEVEL)
      End If
   Next TempRbCdt
   CheckUniqueCollection = True
Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "มีข้อมูล" & TempRbCdt.CUSTOMER_TYPE_NAME & "-" & TempRbCdt.STOCK_GROUP_NAME & "-" & TempRbCdt.REBATE_LEVEL & " ซ้ำ"
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   CheckUniqueCollection = False
End Function
