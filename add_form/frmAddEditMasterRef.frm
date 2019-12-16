VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddEditMasterRef 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditMasterRef.frx":0000
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
      TabIndex        =   8
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Xivess.uctlTextBox txtKeyCode 
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
         Height          =   5805
         Left            =   150
         TabIndex        =   2
         Top             =   1920
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   10239
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
         Column(1)       =   "frmAddEditMasterRef.frx":27A2
         Column(2)       =   "frmAddEditMasterRef.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditMasterRef.frx":290E
         FormatStyle(2)  =   "frmAddEditMasterRef.frx":2A6A
         FormatStyle(3)  =   "frmAddEditMasterRef.frx":2B1A
         FormatStyle(4)  =   "frmAddEditMasterRef.frx":2BCE
         FormatStyle(5)  =   "frmAddEditMasterRef.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditMasterRef.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Xivess.uctlTextBox txtKeyName 
         Height          =   435
         Left            =   1860
         TabIndex        =   1
         Top             =   1320
         Width           =   8325
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin VB.Label lblKeyName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   390
         TabIndex        =   11
         Top             =   1440
         Width           =   1365
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterRef.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   7
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterRef.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   5
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditMasterRef.frx":356A
         ButtonStyle     =   3
      End
      Begin VB.Label lblKeyCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditMasterRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_MasterRef As CMasterRef

Public MasterMode As Long
Public MasterArea As MASTER_TYPE
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_MasterRef.KEY_ID = ID
'      If Not glbDaily.QueryMasterRefDetail(m_MasterRef, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
   End If
   
   If ItemCount > 0 Then
      Call m_MasterRef.PopulateFromRS(1, m_Rs)
      
      txtKeyCode.Text = m_MasterRef.KEY_CODE
      txtKeyName.Text = m_MasterRef.KEY_NAME
      
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

   If ShowMode = SHOW_ADD Then
      If MasterMode = 1 Then
         If Not VerifyAccessRight("MASTER_MAIN_ADD") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 2 Then
              If Not VerifyAccessRight("MASTER_GL_ADD") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 3 Then
        If Not VerifyAccessRight("MASTER_INVENTORY_ADD") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 4 Then
         If Not VerifyAccessRight("MASTER_LEDGER_ADD") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 5 Then
         If Not VerifyAccessRight("MASTER_PRODUCTION_ADD") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 6 Then
   '      If Not VerifyAccessRight("MASTER_PACKAGE_EDIT") Then
   '         Call EnableForm(Me, True)
   '         Exit Function
   '      End If
      ElseIf MasterMode = 7 Then
      ElseIf MasterMode = 8 Then
   
      End If
   Else
      If MasterMode = 1 Then
         If Not VerifyAccessRight("MASTER_MAIN_EDIT") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 2 Then
              If Not VerifyAccessRight("MASTER_GL_EDIT") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 3 Then
         If Not VerifyAccessRight("MASTER_INVENTORY_EDIT", "") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 4 Then
         If Not VerifyAccessRight("MASTER_LEDGER_EDIT") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 5 Then
         If Not VerifyAccessRight("MASTER_PRODUCTION_EDIT") Then
            Call EnableForm(Me, True)
            Exit Function
         End If
      ElseIf MasterMode = 6 Then
'         If Not VerifyAccessRight("MASTER_PACKAGE_EDIT") Then
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
      ElseIf MasterMode = 7 Then
      ElseIf MasterMode = 8 Then
'         If Not VerifyAccessRight("MASTER_PRODUCTION_EDIT") Then
'            Call EnableForm(Me, True)
'            Exit Function
'         End If
      End If
   End If
   
   If Not VerifyTextControl(lblKeyCode, txtKeyCode, False) Then
      Exit Function
   End If

   If Not VerifyTextControl(lblKeyName, txtKeyName, False) Then
      Exit Function
   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_MasterRef.ShowMode = ShowMode
   m_MasterRef.KEY_ID = ID
   m_MasterRef.KEY_CODE = txtKeyCode.Text
   m_MasterRef.KEY_NAME = txtKeyName.Text
   m_MasterRef.MASTER_FLAG = "N"
   m_MasterRef.MASTER_AREA = MasterArea
   
   Call EnableForm(Me, False)
   
'   If Not glbDaily.AddEditMasterRefDetail(m_MasterRef, IsOK, True, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call EnableForm(Me, True)
'      Exit Function
'   End If
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
   
   Set frmAddEditMasterRefItem.ParentForm = Me
   Set frmAddEditMasterRefItem.TempCollection = m_MasterRef.MasterRefDetails
   frmAddEditMasterRefItem.ShowMode = SHOW_ADD
   frmAddEditMasterRefItem.HeaderText = MapText("เพิ่ม")
   Load frmAddEditMasterRefItem
   frmAddEditMasterRefItem.Show 1

   OKClick = frmAddEditMasterRefItem.OKClick

   Unload frmAddEditMasterRefItem
   Set frmAddEditMasterRefItem = Nothing
   
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
   
   If ID1 <= 0 Then
      m_MasterRef.MasterRefDetails.Remove (ID2)
   Else
      m_MasterRef.MasterRefDetails.Item(ID2).Flag = "D"
   End If
   
   Call RefreshGrid
   m_HasModify = True

End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

   If Not VerifyAccessRight("MasterRef_QUERY") Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   Set frmAddEditMasterRefItem.ParentForm = Me
   frmAddEditMasterRefItem.ID = ID
   Set frmAddEditMasterRefItem.TempCollection = m_MasterRef.MasterRefDetails
   frmAddEditMasterRefItem.HeaderText = MapText("แก้ไขสินค้า")
   frmAddEditMasterRefItem.ShowMode = SHOW_EDIT
   Load frmAddEditMasterRefItem
   frmAddEditMasterRefItem.Show 1
   
   OKClick = frmAddEditMasterRefItem.OKClick

   Unload frmAddEditMasterRefItem
   Set frmAddEditMasterRefItem = Nothing
   
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
      ID = m_MasterRef.KEY_ID
      m_MasterRef.QueryFlag = 1
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
         m_MasterRef.QueryFlag = 1
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         m_MasterRef.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call InitGrid1
      
      Call RefreshGrid
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
   
   Set m_MasterRef = Nothing
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
   Col.Width = (ScaleWidth - 500) / 2
   Col.Caption = MapText("กลุ่มลูกค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = (ScaleWidth - 500) / 2
   Col.Caption = MapText("ประเภทลูกค้า")
   
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblKeyCode, MapText("รหัสกลุ่ม"))
   Call InitNormalLabel(lblKeyName, MapText("รายละเอียด"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call txtKeyCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtKeyName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
   Call InitGrid1
   
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
   Set m_MasterRef = New CMasterRef
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_MasterRef.MasterRefDetails Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CMasterRefDetail
   If m_MasterRef.MasterRefDetails.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_MasterRef.MasterRefDetails, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.MASTER_REF_DETAIL_ID
   Values(2) = RealIndex
   Values(3) = CR.MASTER_CUSGROUP_NAME
   Values(4) = CR.MASTER_CUSTYPE_NAME
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub RefreshGrid()
   GridEX1.ItemCount = CountItem(m_MasterRef.MasterRefDetails)
   GridEX1.Rebind
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyCode_Change()
   m_HasModify = True
End Sub
Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

