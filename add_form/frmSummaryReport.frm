VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummaryReport 
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13755
   Icon            =   "frmSummaryReport.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   13755
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   10275
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   18124
      _Version        =   131073
      Begin Threed.SSPanel pnlFooter 
         Height          =   825
         Left            =   0
         TabIndex        =   7
         Top             =   9480
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   1455
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   10140
            TabIndex        =   14
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmSummaryReport.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   11790
            TabIndex        =   15
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdConfig 
            Height          =   525
            Left            =   8520
            TabIndex        =   13
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   8775
         Left            =   7080
         TabIndex        =   8
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   15478
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   4800
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   10
            Top             =   4200
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.ComboBox cboGeneric 
            BeginProperty Font 
               Name            =   "AngsanaUPC"
               Size            =   9
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   630
            Visible         =   0   'False
            Width           =   3855
         End
         Begin PorkShop.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   2400
            TabIndex        =   3
            Top             =   1020
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin PorkShop.uctlDate uctlGenericDate 
            Height          =   435
            Index           =   0
            Left            =   2400
            TabIndex        =   1
            Top             =   240
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin Threed.SSCheck chkCommit 
            Height          =   300
            Index           =   0
            Left            =   2400
            TabIndex        =   4
            Top             =   1560
            Visible         =   0   'False
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   529
            _Version        =   131073
            Caption         =   "SSCheck1"
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   2205
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   735
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   1296
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   4080
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":2ABC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":3398
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":36B4
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   3480
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":3F8E
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin PorkShop.uctlTextBox txtSpace 
            Height          =   435
            Left            =   1230
            TabIndex        =   12
            Top             =   165
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   767
         End
         Begin VB.Label lblSpace 
            Alignment       =   1  'Right Justify
            Caption         =   "Label1"
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Width           =   1155
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   8835
         Left            =   0
         TabIndex        =   0
         Top             =   750
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   15584
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   15.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSummaryReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Public HeaderText As String
Public MasterMode As Long

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_Checks As Collection

Private m_FromDate As Date
Private m_ToDate As Date
Private m_FromRcp As Date
Private m_ToRcp As Date

Private m_MonthID As Long
Private m_YearNo As String
Private TEMP_ROOT_TREE As String

Private C As CReportControl
Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT_EX
   trvMaster.Font.Size = 14
  
  If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MS-1", MapText("รายงานข้อมูลหลัก"), 1, 2)
      Node.Expanded = True
      
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-1", MapText("รายงานข้อมูลลูกค้า"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " MN-2", MapText("รายงานข้อมูลลูกค้า ที่อยู่ ออกจดหมาย"), 1, 2)
      Node.Expanded = True

   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
            
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " S-1", MapText("ระบบงานขาย"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-1-1-2", MapText("ใบส่งสินค้า/ใบกำกับภาษีเป็นชุด"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-21", MapText("รายงานขายเชื่อ เรียงตามวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-22", MapText("รายงานขายเงินสด เรียงตามวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-22-1", MapText("รายงานขายเงินสด/ขายเชื่อ เรียงตามวันที่"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE & " S-1", tvwChild, ROOT_TREE & " S-2-17", MapText("รายงานสรุปยอดขาย แจกแจงเป็นงวด แยกตามลูกค้า สินค้า"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-10", MapText("รายงานรหัสสินค้าและวัตถุดิบ (ST001)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-1-1", MapText("รายงานจำนวนคงคลังและยอดตรวจนับ (ST002)"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2", MapText("รายงานการเคลื่อนไหวคลัง แบบ 1 (ST003)"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2-1", MapText("รายงานการเคลื่อนไหวคลัง แบบ 2 (ST004)"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2-2", MapText("รายงานการเคลื่อนไหวสินค้า แยกตามคลังสินค้า แจกแจงวันที่ (ST004-1)"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3", MapText("รายงานสรุปยอดเคลื่อนไหวสินค้า แยกตามคลังสินค้า (ST005)"), 1, 2)
      Node.Expanded = False


      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-101", MapText("รายงานเอกสารการรับเข้า  (ST101)"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.add(ROOT_TREE, tvwChild, ROOT_TREE & " PD-1", MapText("รายงานการผลิตประจำวัน(PD001)"), 1, 2)
      Node.Expanded = False
   End If
End Sub
Private Sub FillReportInput(R As CReportInterface)
On Error Resume Next
   
   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
         
         If C.Param2 = "MONTH_ID" Then
            m_MonthID = cboGeneric(C.ControlIndex).ListIndex
         End If
         
      End If
      
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
         
         If Len(txtGeneric(C.ControlIndex).Text) = 0 Then
            If C.Param2 = "YEAR_NO" Then
               txtGeneric(C.ControlIndex).Text = Year(Now)
            End If
         End If
         
         If C.Param2 = "YEAR_NO" Then
            m_YearNo = txtGeneric(C.ControlIndex).Text
         End If
         
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_BILL_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_BILL_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               ElseIf C.Param2 = "FROM_RPC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               ElseIf C.Param2 = "TO_RPC_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               End If
            End If
            If C.Param2 = "FROM_BILL_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_BILL_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "FROM_RCP_DATE" Then
               m_FromRcp = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_RCP_DATE" Then
               m_ToRcp = m_Dates(C.ControlIndex).ShowDate
            End If
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
      
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Checks(C.ControlIndex).Value, C.Param2)
         End If
      End If
      
   Next C
End Sub

Private Function VerifyReportInput() As Boolean
On Error Resume Next
   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "D") Then
         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   Next C
   VerifyReportInput = True
End Function
Private Sub cboGeneric_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub chkCommit_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long
Dim ReportMode As Long
   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
      
   ReportKey = trvMaster.SelectedItem.Key
   
   Set Rc = New CReportConfig
   Call Rc.SetFieldValue("REPORT_KEY", ReportKey)
   Call Rc.QueryData(1, m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.GetFieldValue("REPORT_CONFIG_ID")
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If
   
   If ReportKey = "Root MN-2" Then
      ReportMode = 2
   Else
      ReportMode = 1
   End If
   frmReportConfig.ReportMode = ReportMode
   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMaster.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String
Dim ClassName As String
   
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   Key = trvMaster.SelectedItem.Key
   Name = trvMaster.SelectedItem.Text
    
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Key = "Root MN-1" Then
      Set Report = New CReportMain001
      ClassName = "CReportMain001"
   ElseIf Key = "Root MN-2" Then
      Set Report = New CReportMain004
      ClassName = "CReportMain004"
      
      Call Report.AddParam(1, "PREVIEW_TYPE")
   ElseIf Key = "Root MS-1" Then
      Set Report = New CReportMaster001
      ClassName = "CReportMaster001"
      
   ElseIf Key = "Root S-2-21" Then
      Set Report = New CReportBilling021
      ClassName = "CReportBilling021"
   ElseIf Key = "Root S-2-22" Then
      Set Report = New CReportBilling022
      ClassName = "CReportBilling022"
   ElseIf Key = "Root S-2-22-1" Then
      Set Report = New CReportBilling022_1
      ClassName = "CReportBilling022_1"
   ElseIf Key = "Root S-2-17" Then
      Set Report = New CReportBilling018
      ClassName = "CReportBilling018"
   ElseIf Key = "Root S-1-1-2" Then
      Set Report = New CReportNormalRcp001_2
      ClassName = "CReportNormalRcp001_2"
   
   ElseIf Key = "Root 6-10" Then
      Set Report = New CReportInventoryDoc10
      ClassName = "CReportInventoryDoc10"
   ElseIf Key = "Root 6-1-1" Then
      Set Report = New CReportInventoryDoc1_1
      ClassName = "CReportInventoryDoc1_1"
   ElseIf Key = "Root 6-2" Then
      Set Report = New CReportInventoryDoc2
      ClassName = "CReportInventoryDoc2"
   ElseIf Key = "Root 6-2-1" Then
      Set Report = New CReportInventoryDoc2_1
      ClassName = "CReportInventoryDoc2_1"
   ElseIf Key = "Root 6-2-2" Then
      Set Report = New CReportInventoryDoc2_2
      ClassName = "CReportInventoryDoc2_2"
   ElseIf Key = "Root 6-3" Then
      Set Report = New CReportInventoryDoc3
      ClassName = "CReportInventoryDoc3"
   ElseIf Key = "Root 6-101" Then
      Set Report = New CReportInventoryDoc101
      ClassName = "CReportInventoryDoc101"
   ElseIf Key = "Root PD-1" Then
      Set Report = New CReportProduct001
      ClassName = "CReportProduct001"
   End If
   
   SelectFlag = True
   
   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If
      
      Call FillReportInput(Report)
      Call Report.AddParam(Name, "REPORT_NAME")
      Call Report.AddParam(Key, "REPORT_KEY")
      Set frmReport.ReportObject = Report
      frmReport.ClassName = ClassName
      frmReport.Space = Val(txtSpace.Text)
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1
      
      Unload frmReport
      Set frmReport = Nothing
   End If
   
   txtSpace.Text = ""
End Sub

Private Sub Form_Activate()
Dim ItemCount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      m_HasActivate = True
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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
'      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub Form_Resize()
   pnlHeader.Width = ScaleWidth
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   If ScaleWidth <= 0 Then
      trvMaster.Width = 0
   Else
      trvMaster.Width = ScaleWidth - SSFrame2.Width
   End If
   SSFrame2.Left = trvMaster.Width
   If ScaleHeight <= 0 Then
      trvMaster.Height = 0
   Else
      trvMaster.Height = ScaleHeight - pnlHeader.Height - pnlFooter.Height
   End If
   SSFrame2.Height = trvMaster.Height
   pnlFooter.Width = ScaleWidth
   pnlFooter.Top = ScaleHeight - pnlFooter.Height
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20
   cmdConfig.Left = ScaleWidth - cmdExit.Width - 20 - cmdOK.Width - 20 - cmdConfig.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   
   Set m_Rs = Nothing
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_Checks = Nothing
End Sub
Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   SSFrame2.BackColor = GLB_FORM_COLOR
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitNormalLabel(lblSpace, MapText("ระยะห่าง"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   Call InitTreeView
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Call InitFormLayout
   
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
   

   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
   Set m_Checks = New Collection
End Sub

Private Sub UnloadAllControl()
Dim I As Long
Dim j As Long

   I = m_Labels.Count
   While I > 0
      Call Unload(m_Labels(I))
      Call m_Labels.Remove(I)
      I = I - 1
   Wend
   
   I = m_Texts.Count
   While I > 0
      Call Unload(m_Texts(I))
      Call m_Texts.Remove(I)
      I = I - 1
   Wend

   I = m_Dates.Count
   While I > 0
      Call Unload(m_Dates(I))
      Call m_Dates.Remove(I)
      I = I - 1
   Wend

   I = m_Combos.Count
   While I > 0
      Call Unload(m_Combos(I))
      Call m_Combos.Remove(I)
      I = I - 1
   Wend
   
   I = m_TextLookups.Count
   While I > 0
      Call Unload(m_TextLookups(I))
      Call m_TextLookups.Remove(I)
      I = I - 1
   Wend
   
   I = m_Checks.Count
   While I > 0
      Call Unload(m_Checks(I))
      Call m_Checks.Remove(I)
      I = I - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub

Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long


   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "LU") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            If C.OldLine Then
               m_Combos(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Combos(C.ControlIndex).Top = PrevTop - m_Combos(C.ControlIndex - 1).Height
            Else
               m_Combos(C.ControlIndex).Left = PrevLeft
               m_Combos(C.ControlIndex).Top = PrevTop
            End If
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
            If C.OldLine Then
               PrevLeft = m_Combos(C.ControlIndex).Left - CurWidth - 20
            Else
               PrevLeft = m_Combos(C.ControlIndex).Left
            End If
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).Height
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            If C.OldLine Then
               m_Texts(C.ControlIndex).Left = PrevLeft + PrevWidth + 20
               m_Texts(C.ControlIndex).Top = PrevTop - txtGeneric(0).Height
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
               m_Texts(C.ControlIndex).Width = C.Width
            Else
               m_Texts(C.ControlIndex).Left = PrevLeft
               m_Texts(C.ControlIndex).Top = PrevTop
               m_Texts(C.ControlIndex).Width = C.Width
               Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
               m_Texts(C.ControlIndex).Visible = True
                              
               CurTop = PrevTop
               CurLeft = PrevLeft
               CurWidth = PrevWidth
               
               PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).Height
               PrevLeft = m_Texts(C.ControlIndex).Left
               PrevWidth = C.Width
            End If
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).Height
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CH" Then
            m_Checks(C.ControlIndex).Left = PrevLeft
            m_Checks(C.ControlIndex).Top = PrevTop + 100
            m_Checks(C.ControlIndex).Width = C.Width
            m_Checks(C.ControlIndex).Visible = True
         
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
         
            PrevTop = m_Checks(C.ControlIndex).Top + m_Checks(C.ControlIndex).Height
            PrevLeft = m_Checks(C.ControlIndex).Left
            PrevWidth = C.Width
         End If
      
      Else 'Label
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            If C.AllowNull Then
               Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            Else
               Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg, RGB(255, 0, 0))
            End If
            m_Labels(C.ControlIndex).Visible = True
      End If
   Next C
End Sub

Private Sub LoadComboData()
Dim Mr As CMasterRef
   
   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
      
         Set Mr = New CMasterRef
         
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " S-2-17" Then
            If C.ComboLoadID = 1 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitThaiMonth(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-2-2" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 6-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-10" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKGROUP)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_STOCKTYPE)
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 6-101" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_LOCATION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitReport6_2Orderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MS-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMasterTypeName(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitMasterOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " MN-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " MN-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSGROUP)
            ElseIf C.ComboLoadID = 2 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , , , MASTER_CUSTYPE)
            ElseIf C.ComboLoadID = 3 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 4 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         Set Mr = Nothing
      End If
   Next C
   Call EnableForm(Me, True)

End Sub
Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "", Optional KeySearch As String, Optional OldLine As Boolean = False, Optional ToolTipText As String)
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim ChIdx As Long

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   ChIdx = m_Checks.Count + 1
   
   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
      lblGeneric(LblIdx).ToolTipText = ToolTipText
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
      C.OldLine = OldLine
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
      C.OldLine = OldLine
      txtGeneric(TxtIdx).SetKeySearch (KeySearch)
      
      If Param1 = "YEAR_NO" Then
         If Len(m_YearNo) > 0 Then
            txtGeneric(TxtIdx).Text = m_YearNo
         Else
            txtGeneric(TxtIdx).Text = Year(Now) + 543
         End If
      End If
      
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx
      
      
      If DateIdx = 1 Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      ElseIf DateIdx = 2 Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_ToDate
         End If
      ElseIf DateIdx = 3 Then
         If m_FromRcp > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromRcp
         Else
            Call GetFirstLastDate(Now, m_FromRcp, m_ToRcp)
            uctlGenericDate(DateIdx).ShowDate = m_FromRcp
         End If
      ElseIf DateIdx = 4 Then
         If m_ToRcp > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_ToRcp
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToRcp)
            uctlGenericDate(DateIdx).ShowDate = m_ToRcp
         End If
      ElseIf DateIdx = 5 Then
         If m_FromDate > 0 Then
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         Else
            Call GetFirstLastDate(Now, m_FromDate, m_ToDate)
            uctlGenericDate(DateIdx).ShowDate = m_FromDate
         End If
      End If
      
   ElseIf ControlType = "LU" Then
'         Load uctlTextLookup(LkupIdx)
'         Call m_TextLookups.Add(uctlTextLookup(LkupIdx))
'         C.ControlIndex = LkupIdx
   ElseIf ControlType = "CH" Then
      Load chkCommit(ChIdx)
      Call m_Checks.add(chkCommit(ChIdx))
      Call InitCheckBox(chkCommit(ChIdx), TextMsg)
      C.ControlIndex = ChIdx
   End If
   
   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.add(C)
   Set C = Nothing
End Sub

Private Sub InitReport1_1()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_2()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "GROUP_ID", "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_3()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).WIDTH, True, "", 1, "GROUP_ID", "GROUP_NAME")
'   Call LoadControl("L", lblGeneric(0).WIDTH, True, GetTextMessage("TEXT-KEY71"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean
   
'   If LastKey = Node.Key Then
'      Exit Sub
'   End If
   
   LastKey = Node.Key
   
   Status = True
   QueryFlag = False
   
   Call UnloadAllControl
   
   If Node.Children > 0 Then
      cmdOK.Enabled = False
      Exit Sub
   End If
   
   If MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MAIN_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("PRODUCT_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 5 Then
      If Not VerifyAccessRight("LEDGER_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("INVENTORY_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("TAGET_REPORT_" & Node.Text, Node.Text) Then
         Call EnableForm(Me, True)
         cmdOK.Enabled = False
         Exit Sub
      End If
   End If
   
   cmdOK.Enabled = True
   
   If Node.Key = ROOT_TREE & " MS-1" Then
      Call InitReportMS_1
  ElseIf Node.Key = ROOT_TREE & " MS-2" Then
      Call InitReportMS_2
  ElseIf Node.Key = ROOT_TREE & " MS-3" Then
      Call InitReportMS_3
   ElseIf Node.Key = ROOT_TREE & " S-1-1-2" Then
      Call InitReportS_1_1
   ElseIf Node.Key = ROOT_TREE & " S-2-17" Then
      Call InitReportS_2_17
   ElseIf Node.Key = ROOT_TREE & " S-2-21" Then
      Call InitReportS_2_21
   ElseIf Node.Key = ROOT_TREE & " S-2-22" Then
      Call InitReportS_2_21
   ElseIf Node.Key = ROOT_TREE & " S-2-22-1" Then
      Call InitReportS_2_21
   ElseIf Node.Key = ROOT_TREE & " 6-1-1" Then
      Call InitReport6_1
   ElseIf Node.Key = ROOT_TREE & " 6-2" Then
      Call InitReport6_2
   ElseIf Node.Key = ROOT_TREE & " 6-2-1" Then
      Call InitReport6_2_1
   ElseIf Node.Key = ROOT_TREE & " 6-2-2" Then
      Call InitReport6_2_2
   ElseIf Node.Key = ROOT_TREE & " 6-3" Then
      Call InitReport6_3
   ElseIf Node.Key = ROOT_TREE & " 6-10" Then
      Call InitReport6_10
   ElseIf Node.Key = ROOT_TREE & " 6-101" Then
      Call InitReport6_101
   ElseIf Node.Key = ROOT_TREE & " MN-1" Then
      Call InitReportMN_1
   ElseIf Node.Key = ROOT_TREE & " MN-2" Then
      Call InitReportMN_2
   ElseIf Node.Key = ROOT_TREE & " PD-1" Then
      Call InitReportPD_1
  End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
   End If
   Call EnableForm(Me, True)
End Sub
Private Sub InitReport6_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ณ วันที่"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, False, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, False, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_2_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_2_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport6_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงหน่วย", , "SHOW_UNIT_NAME")
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport6_101()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "LOCATION_ID", "LOCATION_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("คลัง"))
   
   '8 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "DECIMAL_AMOUNT")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนทศนิยม"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportS_2_17()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
         
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "สรุป", , "SHOW_SUMMARY")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงจำนวน", , "NOT_SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่แสดงมูลค่า", , "NOT_SHOW_PRICE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "ไม่รวมส่วนลด", , "NOT_DISCOUNT")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_31()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   If TEMP_ROOT_TREE = " S-2-31" Then
      Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   End If
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReport6_10()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_STOCK_NO", , "STOCK_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสวัตถุดิบ"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_STOCK_NO", , "STOCK_NO", True)
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "STOCK_GROUP", "STOCK_GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มสินค้า/วัตถุดิบ"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "STOCK_TYPE", "STOCK_TYPE_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทสินค้า/วัตถุดิบ"))
   
   '9 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
      
   Call LoadControl("CH", cboGeneric(0).Width, True, "ยกเลิก", , "EXCEPTION_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงค่าใช้จ่าย", , "EXPENSE_FLAG")
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportP_1_3()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("วันที่ส่งของ"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "UNIT_CHANGE1", "UNIT_CHANGE_NAME1")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย1"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "UNIT_CHANGE2", "UNIT_CHANGE_NAME2")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย2"))
   
   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "UNIT_CHANGE3", "UNIT_CHANGE_NAME3")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย3"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "UNIT_CHANGE4", "UNIT_CHANGE_NAME4")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย4"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 5, "UNIT_CHANGE5", "UNIT_CHANGE_NAME5")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย5"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 6, "UNIT_CHANGE6", "UNIT_CHANGE_NAME6")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย6"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 7, "UNIT_CHANGE7", "UNIT_CHANGE_NAME7")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย7"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 8, "UNIT_CHANGE8", "UNIT_CHANGE_NAME8")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย8"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 9, "UNIT_CHANGE9", "UNIT_CHANGE_NAME9")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย9"))
   
   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 10, "UNIT_CHANGE10", "UNIT_CHANGE_NAME10")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รวมหน่วย10"))
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMS_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "MASTER_AREA", "MASTER_AREA_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หัวข้อข้อมูลหลัก"))
      
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMN_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มลูกหนี้"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMN_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มลูกหนี้"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "AMOUNT_PER_ITEM")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนต่อลูกค้า(Min:1)"))
      
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "AMOUNT_ENTERPRISE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จำนวนต่อบริษัท(Min:0)"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FONT_SIZE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("Font (MAX:16)"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReportMS_2()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
      
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "APAR_GROUP")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("กลุ่มลูกหนี้"))
   
   '7 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "APAR_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกหนี้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReportMS_3() 'พนักงานขาย sale
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
      
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub cboGeneric_Click(Index As Integer)
Dim Node As Node
Dim TempID As Long

   Set Node = trvMaster.SelectedItem
   
   If (Node.Key = ROOT_TREE & " MN-1" Or Node.Key = ROOT_TREE & " MN-2") Then
      If Index = 1 Then
         TempID = cboGeneric(Index).ItemData(Minus2Zero(cboGeneric(Index).ListIndex))
         If TempID > 0 Then
            Call LoadMaster(cboGeneric(Index + 1), , , , MASTER_CUSTYPE, , TempID)
         End If
      End If
   End If
End Sub
Private Sub InitReportS_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "DOCUMENT_NO_SEARCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หมายเลขเอกสาร"), , , , , , "ตัวอย่าง ขายสด HS,ใบส่งของ IV55,ใบกำกับ IVV5509,PO....")
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportSL_1_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
      
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 1, "MONTH_ID", "MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เป้าการขายเดือน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เป้าการขายปี"))
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่เปรียบเทียบ"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE_EX")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่เปรียบเทียบ"))
      
   '4 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "จำนวน", , "SHOW_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ยอดขาย", , "SHOW_PRICE")
   Call LoadControl("CH", cboGeneric(0).Width, True, "จำนวนคืน", , "SHOW_RETURN_AMOUNT")
   Call LoadControl("CH", cboGeneric(0).Width, True, "ยอดคืน", , "SHOW_RETURN_PRICE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportP_1_5()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '1 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "TRANSPORTOR1", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง1"))
      
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR1_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รายจ่ายขนส่ง1"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "TRANSPORTOR2", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง2"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR2_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รายจ่ายขนส่ง2"))
   
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "TRANSPORTOR3", "")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ขนส่ง3"))
   
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "TRANSPORTOR3_PRICE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รายจ่ายขนส่ง3"))
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "รวมรายการของแถม", , "INCLUDE_FREE")
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "คิดยอดจากใบ PO", , "PO_FLAG")
   Call LoadControl("CH", cboGeneric(0).Width, True, "คิดยอดจากใบส่งของและขายสด", , "INVOICE_FLAG")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportS_2_21()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่บิล"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, False, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่บิล"))
   
   '3 =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_APAR_CODE", , "CUSTOMER_CODE")
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_APAR_CODE", , "CUSTOMER_CODE", True)
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SALE_CODE", , "SALE_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงานขาย"))
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SALE_CODE", , "SALE_CODE", True)
   
   Call LoadControl("CH", cboGeneric(0).Width, True, "แสดงรายละเอียด", , "SHOW_DETAIL")
   
   Call ShowControl
   Call LoadComboData
   
End Sub
Private Sub InitReportPD_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long
   
   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
 '1 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_BILL_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   
   
   Call ShowControl
   Call LoadComboData
   
End Sub

