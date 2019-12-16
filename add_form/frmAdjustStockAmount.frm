VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjustStockAmount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmAdjustStockAmount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   7035
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   12409
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   5520
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   8
         Top             =   5880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   900
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtFromStockNO 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin PorkShop.uctlTextBox txtToStockNo 
         Height          =   465
         Left            =   4980
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   2835
         Left            =   1860
         TabIndex        =   3
         Top             =   1920
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   5001
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin PorkShop.uctlTextBox txtAdjustToAmount 
         Height          =   465
         Left            =   1860
         TabIndex        =   4
         Top             =   4920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin Threed.SSCheck chkCheckAll 
         Height          =   435
         Left            =   8640
         TabIndex        =   18
         Top             =   1320
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblAdjustToAmount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   5040
         Width           =   1605
      End
      Begin VB.Label lblFromStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label lblToStockNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   5
         Top             =   6180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAdjustStockAmount.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   13
         Top             =   6000
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   11
         Top             =   6000
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   6
         Top             =   6180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAdjustStockAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public HeaderText As String
Private m_LocationColls As Collection
Private Sub InitTreeView()
Dim Node As Node
   trvMain.Font.Size = 17
   trvMain.Font.Name = "JasmineUPC"
   
   ClearTreeView (trvMain.hwnd)
   Call GenerateTree1(m_LocationColls, Nothing, "ROOT", 0, 0)
End Sub
Private Sub GenerateTree1(Locations As Collection, N As Node, NodeID As String, PID As Long, Level As Long)
Dim O As CMasterRef
Dim Node As Node
Dim NewNodeID As String
Dim L As Long

   For Each O In Locations
      Set Node = trvMain.Nodes.add(, tvwFirst, NodeID & O.KEY_ID, O.KEY_NAME & " (" & O.KEY_CODE & ")")
      Node.Tag = O.KEY_ID
      Node.Checked = True
      O.Flag = "Y"
      Node.Expanded = False
   Next O
End Sub
Private Sub GenerateTree2(Locations As Collection, N As Node, NodeID As String, PID As Long, Level As Long)
Dim O As CMasterRef
Dim Node As Node
Dim NewNodeID As String
Dim L As Long

   For Each O In Locations
      Set Node = trvMain.Nodes.add(, tvwFirst, NodeID & O.KEY_ID, O.KEY_NAME & " (" & O.KEY_CODE & ")")
      Node.Tag = O.KEY_ID
      O.Flag = "N"
      Node.Expanded = False
   Next O
End Sub
Private Sub UpdateCollection(ID As Long, Flag As Boolean)
Dim D As CMasterRef
   For Each D In m_LocationColls
      If D.KEY_ID = ID Then
         If Flag Then
            D.Flag = "Y"
         Else
            D.Flag = "N"
         End If
      End If
   Next D
End Sub
Private Sub chkCheckAll_Click(Value As Integer)
   ClearTreeView (trvMain.hwnd)
   If chkCheckAll.Value = ssCBChecked Then
      Call GenerateTree1(m_LocationColls, Nothing, "ROOT", 0, 0)
   Else
      Call GenerateTree2(m_LocationColls, Nothing, "ROOT", 0, 0)
   End If
End Sub

Private Sub trvMain_NodeCheck(ByVal Node As MSComctlLib.Node)
   Call UpdateCollection(Node.Tag, Node.Checked)
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
Dim IsOK As Boolean
   Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   Status = AdjustStockCode
   
   Me.Enabled = True
   
   If Status Then
      Call glbDaily.CommitTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดดเสร็จสมบูรณ์"
      glbErrorLog.ShowUserError
   Else
      Call glbDaily.RollbackTransaction
      glbErrorLog.LocalErrorMsg = "การอัฟเดด ERROR"
      glbErrorLog.ShowUserError
   End If
   
   OKClick = True
   Unload Me
   Exit Sub
   
End Sub
Private Sub Form_Activate()
      Me.Refresh
      DoEvents
      
      Call LoadMaster(Nothing, m_LocationColls, , , MASTER_LOCATION)
      
      Call InitTreeView
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   End If

End Sub
Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblToDate, "ถึงวันที่", RGB(255, 0, 0))
   Call InitNormalLabel(lblFromStockNo, "จากรหัส Stock")
   Call InitNormalLabel(lblToStockNo, "ถึงรหัส Stock")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblAdjustToAmount, "ยอดปรับ")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call InitCheckBox(chkCheckAll, "เลือกทั้งหมด")
   
   chkCheckAll.Value = ssCBChecked
   
   Call txtFromStockNO.SetKeySearch("STOCK_NO")
   Call txtToStockNo.SetKeySearch("STOCK_NO")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   
   Set m_LocationColls = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Function AdjustStockCode() As Boolean
Dim m_LotItem As CLotItem
Dim m_Rs  As ADODB.Recordset
Dim ItemCount As Long
Dim I As Long
Dim IsOK As Boolean
Dim Mr As CMasterRef
Dim Pi As CStockCode
Dim TempPartColl As Collection

Dim Ivd As CInventoryDoc
Dim TempLotItem As CLotItem

   Set TempPartColl = New Collection
   Set m_Rs = New ADODB.Recordset
   AdjustStockCode = False
      
   Call LoadStockCode(Nothing, TempPartColl)
   
   Set m_LotItem = New CLotItem
   m_LotItem.TO_DOC_DATE = uctlToDate.ShowDate
   m_LotItem.FROM_STOCK_NO = txtFromStockNO.Text
   m_LotItem.TO_STOCK_NO = txtToStockNo.Text
   Call m_LotItem.QueryData(50, m_Rs, ItemCount, False)
   
   I = 0
   prgProgress.Min = 0
   If m_Rs.RecordCount > 0 Then
      prgProgress.Max = m_Rs.RecordCount
   End If
   
   Set Ivd = New CInventoryDoc
   Ivd.ShowMode = SHOW_ADD
   Ivd.DOCUMENT_NO = "***ADJUST_" & uctlToDate.ShowDate & "_" & txtFromStockNO.Text & "-" & txtToStockNo.Text
   Ivd.DOCUMENT_DATE = uctlToDate.ShowDate
   Ivd.DOCUMENT_TYPE = ADJUST_DOCTYPE
   Ivd.DOCUMENT_DESC = "ตั้งยอดเป็น " & Val(txtAdjustToAmount.Text)
   Ivd.CANCEL_FLAG = "N"
      
   While Not m_Rs.EOF
      Call m_LotItem.PopulateFromRS(50, m_Rs)
      
      I = I + 1
      prgProgress.Value = I
      txtPercent.Text = MyDiffEx(I, m_Rs.RecordCount) * 100
      Me.Refresh
      
      Set Mr = GetObject("CMasterRef", m_LocationColls, Trim(Str(m_LotItem.LOCATION_ID)))
      
      If (Round(m_LotItem.SUM_AMOUNT, 2) <> (Val(txtAdjustToAmount.Text))) And (Mr.Flag = "Y") Then
         Set TempLotItem = New CLotItem
         
         TempLotItem.Flag = "A"
         TempLotItem.PART_ITEM_ID = m_LotItem.PART_ITEM_ID
         TempLotItem.LOCATION_ID = m_LotItem.LOCATION_ID
         
         If m_LotItem.SUM_AMOUNT > Val(txtAdjustToAmount.Text) Then 'ตอนนี้เป็น + เลยต้องการเบิกออก
            TempLotItem.TX_AMOUNT = m_LotItem.SUM_AMOUNT - Val(txtAdjustToAmount.Text)
            TempLotItem.MULTIPLIER = -1
            TempLotItem.TX_TYPE = "E"
         Else
            TempLotItem.TX_AMOUNT = Val(txtAdjustToAmount.Text) - m_LotItem.SUM_AMOUNT
            TempLotItem.MULTIPLIER = 1
            TempLotItem.TX_TYPE = "I"
         End If
         
         Set Pi = GetObject("CStockCode", TempPartColl, Trim(Str(m_LotItem.PART_ITEM_ID)))
         TempLotItem.UNIT_TRAN_ID = Pi.UNIT_CHANGE_ID
         TempLotItem.UNIT_MULTIPLE = 1
         
         Call Ivd.ImportExportItems.add(TempLotItem)
         
         Set TempLotItem = Nothing
      End If
      m_Rs.MoveNext
   Wend
   
   If Ivd.ImportExportItems.Count > 0 Then
      If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Exit Function
      End If
   End If
   
   Set Ivd = Nothing
   Set TempPartColl = Nothing
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_LotItem = Nothing
   AdjustStockCode = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set m_LocationColls = Nothing
End Sub
