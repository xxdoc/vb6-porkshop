VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAddEditJob1 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditJob1.frx":0000
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
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   6
         Top             =   2355
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
      Begin PorkShop.uctlTextBox txtJobNo 
         Height          =   435
         Left            =   2400
         TabIndex        =   0
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
         Height          =   4845
         Left            =   150
         TabIndex        =   7
         Top             =   2880
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8546
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
         Column(1)       =   "frmAddEditJob1.frx":27A2
         Column(2)       =   "frmAddEditJob1.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditJob1.frx":290E
         FormatStyle(2)  =   "frmAddEditJob1.frx":2A6A
         FormatStyle(3)  =   "frmAddEditJob1.frx":2B1A
         FormatStyle(4)  =   "frmAddEditJob1.frx":2BCE
         FormatStyle(5)  =   "frmAddEditJob1.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditJob1.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtJobDesc 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   1320
         Width           =   8325
         _ExtentX        =   12568
         _ExtentY        =   767
      End
      Begin PorkShop.uctlDate uctlJobDate 
         Height          =   405
         Left            =   6360
         TabIndex        =   1
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtFormulaAmount 
         Height          =   435
         Left            =   10440
         TabIndex        =   4
         Top             =   1800
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   767
      End
      Begin PorkShop.uctlTextLookup uctlFormula 
         Height          =   435
         Left            =   1860
         TabIndex        =   3
         Top             =   1800
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   767
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   2640
         TabIndex        =   18
         Top             =   7830
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   1860
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   860
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":3250
         ButtonStyle     =   3
      End
      Begin VB.Label lblFormulaAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8640
         TabIndex        =   16
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblFormula 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label lblJobDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   390
         TabIndex        =   13
         Top             =   1440
         Width           =   1365
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   9
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
         TabIndex        =   5
         Top             =   7830
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditJob1.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblJobNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmAddEditJob1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Job As CJob

Private m_Formulas  As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

Private m_Cd As Collection
Private DocAdd As Long
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Job.JOB_ID = ID
      If Not glbDaily.QueryJob(m_Job, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_Job.PopulateFromRS(1, m_Rs)
      
      txtJobNo.Text = m_Job.JOB_NO
      txtJobDesc.Text = m_Job.JOB_DESC
      uctlJobDate.ShowDate = m_Job.JOB_DATE
      
      txtFormulaAmount.Text = m_Job.FORMULA_AMOUNT
      uctlFormula.MyCombo.ListIndex = IDToListIndex(uctlFormula.MyCombo, m_Job.FORMULA_ID)
      
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If m_Job.CollJobInputs.Count > 0 Or m_Job.CollJobOutputs.Count > 0 Then
      txtFormulaAmount.Enabled = False
      uctlFormula.Enabled = False
      cmdAdd.Enabled = False
      cmdClear.Enabled = True
   Else
      txtFormulaAmount.Enabled = True
      uctlFormula.Enabled = True
      cmdAdd.Enabled = True
      cmdClear.Enabled = False
   End If
   
   Call TabStrip1_Click
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc
   
   If Not VerifyTextControl(lblJobNo, txtJobNo, False) Then
      Exit Function
   End If
   
   If Not VerifyDate(lblJobDate, uctlJobDate, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblFormula, uctlFormula.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblFormulaAmount, txtFormulaAmount, False) Then
      Exit Function
   End If
   
   If Not VerifyLockDate(uctlJobDate.ShowDate, m_Job.JOB_DATE) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not VerifyLockInventoryDate(uctlJobDate.ShowDate, m_Job.JOB_DATE) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
      
   If Not CheckUniqueNs(JOB_NO, txtJobNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtJobNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      DocAdd = DocAdd + 1
      Call cmdAuto_Click
      Exit Function
   End If
      
   m_Job.AddEditMode = ShowMode
   m_Job.JOB_ID = ID
   m_Job.JOB_NO = txtJobNo.Text
   m_Job.JOB_DESC = txtJobDesc.Text
   m_Job.JOB_DATE = uctlJobDate.ShowDate
   m_Job.FORMULA_ID = uctlFormula.MyCombo.ItemData(Minus2Zero(uctlFormula.MyCombo.ListIndex))
   m_Job.FORMULA_AMOUNT = txtFormulaAmount.Text
   
   Call EnableForm(Me, False)
   
   Call glbDaily.Job2InventoryDoc(m_Job, Ivd, 1000)             'ใบสั่งผลิต
   
   Call glbDaily.StartTransaction
   If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   m_Job.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   
   If Not glbDaily.AddEditJob(m_Job, IsOK, False, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call glbDaily.RollbackTransaction
      Call EnableForm(Me, True)
      Exit Function
   End If
   
   Call glbDaily.CommitTransaction
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cmdAdd_Click()
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyCombo(lblFormula, uctlFormula.MyCombo, False) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblFormulaAmount, txtFormulaAmount, False) Then
      Exit Sub
   End If
   
   Dim TempJobItem As CJobItem
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim IsOK As Boolean
   Dim TempCount As Long
   Dim TempFormula As CFormula
   Dim TempRs As ADODB.Recordset
   Dim TempFormulaItem As CFormulaItem
   Dim I  As Long
   Dim TempCollection As Collection
   
   IsOK = True
   Set TempFormula = New CFormula
   Set TempRs = New ADODB.Recordset
   
   TempFormula.FORMULA_ID = uctlFormula.MyCombo.ItemData(Minus2Zero(uctlFormula.MyCombo.ListIndex))
   TempFormula.QueryFlag = 1
   
   If Not glbDaily.QueryFormula(TempFormula, TempRs, TempCount, IsOK, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Sub
   End If
   
   For I = 1 To 2
      If I = 1 Then
         Set TempCollection = TempFormula.CollFormulaInputs
      Else
         Set TempCollection = TempFormula.CollFormulaOutputs
      End If
      For Each TempFormulaItem In TempCollection
         Set TempJobItem = New CJobItem
         TempJobItem.Flag = "A"
         TempJobItem.PART_ITEM_ID = TempFormulaItem.PART_ITEM_ID
         TempJobItem.PART_NO = TempFormulaItem.PART_NO
         TempJobItem.PART_DESC = TempFormulaItem.PART_DESC
         
         TempJobItem.LOCATION_ID = TempFormulaItem.LOCATION_ID
         TempJobItem.LOCATION_NO = TempFormulaItem.LOCATION_NO
         TempJobItem.LOCATION_NAME = TempFormulaItem.LOCATION_NAME
         
         TempJobItem.TX_TYPE = TempFormulaItem.TX_TYPE
         
         TempJobItem.TX_AMOUNT = TempFormulaItem.TX_AMOUNT * Val(txtFormulaAmount.Text)
         
         TempJobItem.UNIT_TRAN_ID = TempFormulaItem.UNIT_TRAN_ID
         TempJobItem.UNIT_MULTIPLE = TempFormulaItem.UNIT_MULTIPLE
         TempJobItem.UNIT_TRAN_NAME = TempFormulaItem.UNIT_TRAN_NAME
         TempJobItem.UNIT_NAME = TempFormulaItem.UNIT_NAME
         
         If I = 1 Then
            Call m_Job.CollJobInputs.add(TempJobItem)
         Else
            Call m_Job.CollJobOutputs.add(TempJobItem)
         End If
         
         Set TempJobItem = Nothing
      Next TempFormulaItem
   Next I
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Call PopulateGuiID(m_Job)
   
   Call RefreshGrid(1)
   
   Set TempFormula = Nothing
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   uctlFormula.Enabled = False
   txtFormulaAmount.Enabled = False
   cmdAdd.Enabled = False
   m_HasModify = True
End Sub

Private Sub cmdClear_Click()
Dim TempJobItem As CJobItem
   For Each TempJobItem In m_Job.CollJobInputs
      TempJobItem.Flag = "D"
   Next TempJobItem
   
   For Each TempJobItem In m_Job.CollJobOutputs
      TempJobItem.Flag = "D"
   Next TempJobItem
         
   m_HasModify = True
      
   Call SaveData
   
   ShowMode = SHOW_EDIT
   ID = m_Job.JOB_ID
   m_Job.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
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
      ID = m_Job.JOB_ID
      m_Job.QueryFlag = 1
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
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
      Call LoadFormula(uctlFormula.MyCombo, m_Formulas)
      Set uctlFormula.MyCollection = m_Formulas
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Job.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Job.QueryFlag = 0
         Call QueryData(False)
         Call cmdAuto_Click
         uctlJobDate.ShowDate = Now
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
   
   Set m_Job = Nothing
   Set m_Formulas = Nothing
   Set m_Cd = Nothing
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
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 2000
   Col.Caption = MapText("รหัสวัตถุดิบ")

   Set Col = GridEX1.Columns.add '4
   Col.Width = ScaleWidth - 9500
   Col.Caption = MapText("ชื่อวัตถุดิบ")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน/สูตร")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("หน่วย")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3000
   Col.Caption = MapText("สถานที่จัดเก็บ")
   
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
   Col.Width = 2000
   Col.Caption = MapText("รหัสสินค้า")

   Set Col = GridEX1.Columns.add '4
   Col.Width = ScaleWidth - 9500
   Col.Caption = MapText("ชื่อสินค้า")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("จำนวน/สูตร")
   
   Set Col = GridEX1.Columns.add '5
   Col.Width = 2000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("หน่วย")
   
   Set Col = GridEX1.Columns.add '3
   Col.Width = 3000
   Col.Caption = MapText("สถานที่จัดเก็บ")
   
End Sub

Private Sub InitFormLayout()

   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblJobNo, MapText("หมายเลข JOB"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblJobDate, MapText("วันที่"))
   Call InitNormalLabel(lblFormula, MapText("สูตร"))
   Call InitNormalLabel(lblFormulaAmount, MapText("ยอดสูตร"))
   
   Call txtJobNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtFormulaAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่มจากสูตร (F7)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   Call InitMainButton(cmdClear, MapText("CLEAR"))
   
   Call InitGrid1
   Call InitGrid2
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.add().Caption = MapText("                       วัตถุดิบ                     ")
   TabStrip1.Tabs.add().Caption = MapText("              ผลิตภัณฑ์/สินค้า              ")
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
   Set m_Job = New CJob
   Set m_Formulas = New Collection
   Set m_Cd = New Collection
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long
Dim CR As CJobItem
   
   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Or TabStrip1.SelectedItem.Index = 2 Then
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

      Values(1) = CR.JOB_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.PART_NO
      Values(4) = CR.PART_DESC
      Values(5) = FormatNumber(MyDiff(CR.TX_AMOUNT, CR.UNIT_MULTIPLE))
      Values(6) = CR.UNIT_TRAN_NAME & " X " & CR.UNIT_MULTIPLE & " " & CR.UNIT_NAME
      Values(7) = CR.LOCATION_NAME
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
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
      GridEX1.ItemCount = CountItem(m_Job.CollJobInputs)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      GridEX1.ItemCount = CountItem(m_Job.CollJobOutputs)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
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

Private Function GetCollection(Index As Long) As Collection
   If Index = 1 Then
      Set GetCollection = m_Job.CollJobInputs
   ElseIf Index = 2 Then
      Set GetCollection = m_Job.CollJobOutputs
   End If
End Function

Private Sub txtFormulaAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtJobDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtJobNo_Change()
   m_HasModify = True
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
   cmdClear.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

Private Sub uctlFormula_Change()
   m_HasModify = True
End Sub

Private Sub uctlJobDate_HasChange()
   m_HasModify = True
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
Private Function GetNextGuiID(BD As CJob) As Long
Dim Di As CJobItem
Dim MaxId As Long

   MaxId = 0
   
   For Each Di In BD.CollJobInputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   For Each Di In BD.CollJobOutputs
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di
   
   GetNextGuiID = MaxId + 1
End Function
Private Sub PopulateGuiID(BD As CJob)
Dim Di As CJobItem
   
   For Each Di In BD.CollJobInputs
      If Di.LOCATION_ID > 0 Then
         If Di.Flag = "A" Then
            Di.LINK_ID = GetNextGuiID(BD)
         End If
      End If
   Next Di
   
   For Each Di In BD.CollJobOutputs
      If Di.LOCATION_ID > 0 Then
         If Di.Flag = "A" Then
            Di.LINK_ID = GetNextGuiID(BD)
         End If
      End If
   Next Di
   
End Sub
Private Sub cmdAuto_Click()
Dim ID As Long
Dim Cd As CConfigDoc
Dim TempStr As String
Dim I As Long
   
   If Len(txtJobNo.Text) > 0 Then
      SendKeys ("{TAB}")
      Exit Sub
   End If
   
   ID = 1000 'บังคับไปเลยว่าเป็น 1000 ซึ่งเป็นไปสั่งผลิต
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         txtJobNo.Text = Cd.GetFieldValue("PREFIX")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         txtJobNo.Text = txtJobNo.Text & Format(Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd, TempStr)
         m_Job.RUNNING_NO = Cd.GetFieldValue("RUNNING_NO") + 1 + DocAdd
         m_Job.CONFIG_DOC_TYPE = ID
         
         Call txtJobNo.SetSelectText(Len(txtJobNo.Text) - Cd.GetFieldValue("DIGIT_AMOUNT"), Cd.GetFieldValue("DIGIT_AMOUNT"))
      Else
         txtJobNo.Text = ""
      End If
   End If
End Sub

