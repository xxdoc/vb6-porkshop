VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommissionChart 
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmCommissionChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   7905
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   13944
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCommissionChart.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCommissionChart.frx":11A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   5415
         Left            =   0
         TabIndex        =   4
         Top             =   1560
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   9551
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
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
      Begin Rebate.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1740
         TabIndex        =   2
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Rebate.uctlTextBox txtMasterFromToNo 
         Height          =   435
         Left            =   1740
         TabIndex        =   0
         Top             =   120
         Width           =   2385
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextBox txtMasterFromToDesc 
         Height          =   435
         Left            =   1740
         TabIndex        =   1
         Top             =   600
         Width           =   8325
         _ExtentX        =   5001
         _ExtentY        =   767
      End
      Begin Rebate.uctlDate uctlToDate 
         Height          =   405
         Left            =   7800
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8520
         TabIndex        =   8
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionChart.frx":1A7E
         ButtonStyle     =   3
      End
      Begin VB.Label lblMasterFromToNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   450
         TabIndex        =   13
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblMasterFromToDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   270
         TabIndex        =   12
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6480
         TabIndex        =   11
         Top             =   1140
         Width           =   1155
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3330
         TabIndex        =   7
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionChart.frx":1D98
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   60
         TabIndex        =   5
         Top             =   7200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCommissionChart.frx":20B2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1680
         TabIndex        =   6
         Top             =   7200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   525
         Left            =   10200
         TabIndex        =   9
         Top             =   7200
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   705
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1244
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
End
Attribute VB_Name = "frmCommissionChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmCommissionChart"
Private Const ROOT_TREE = "R"
Private HasActivate As Boolean
Private m_HasModify As Boolean
Private m_MasterFromTo As CMasterFromTo
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public DocumentType As MASTER_COMMISSION_AREA

Private m_Commissions As Collection
Private Function GetCommissionChart(TempID As Long, TempCol As Collection) As CCommissionChart
Dim L As CCommissionChart

   For Each L In TempCol
      If L.GetFieldValue("COMMISSION_CHART_ID") = TempID Then
         Set GetCommissionChart = L
         Exit Function
      End If
   Next L
End Function

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim TableName As String
Dim L As CCommissionChart

      
   If trvMain.SelectedItem Is Nothing Then
'      glbErrorLog.LocalErrorMsg = GetTextMessage("TEXT-KEY569")
'      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If trvMain.Nodes.Count <= 0 Then
      Exit Sub
   End If
   
   ID = Val(trvMain.SelectedItem.Tag)
   
   Set L = GetObject("CCommissionChart", m_Commissions, Trim(Str(ID)))
   
   glbErrorLog.LocalErrorMsg = "ต้องการลบข้อมูล " & L.GetFieldValue("EMP_NAME") & " " & L.GetFieldValue("EMP_LNAME") & " ใช่หรือไม่ ?"
   If glbErrorLog.AskMessage = vbNo Then
      Exit Sub
   End If
   
   Call L.SetFieldValue("COMMISSION_CHART_ID", ID)
   
   If Not glbDaily.DeleteCommissionChart(L, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Call LoadCommissionChart(Nothing, m_Commissions, m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID"))
   
   
   Call InitMainTreeview("", m_Commissions)
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Activate()
On Error GoTo ErrorHandler
Dim ItemCount As Long

   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Activate"

   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      
         Call LoadCommissionChart(Nothing, m_Commissions, m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID"))
         Call InitMainTreeview("", m_Commissions)
         m_HasModify = False
      End If
   End If
   
   
   Call EnableForm(Me, True)
   Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
     glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
   ElseIf Shift = 0 And KeyCode = 116 Then
      'Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
      'Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
      'Call cmdPrint_Click
   ElseIf Shift = 0 And KeyCode = 27 Then
      Call cmdExit_Click
   End If
End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long

   If m_HasModify Or (m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID") <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   
   Call EnableForm(Me, False)

      
   frmAddEditCommissionChart.HeaderText = "เพิ่มข้อมูลแผนภูมิการคิดคอมมิตชั่น"
   frmAddEditCommissionChart.ShowMode = SHOW_ADD
   frmAddEditCommissionChart.FK_ID = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID")
   Load frmAddEditCommissionChart
   frmAddEditCommissionChart.Show 1

   If frmAddEditCommissionChart.OKClick Then
      Call EnableForm(Me, False)
      Call LoadCommissionChart(Nothing, m_Commissions, m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID"))
      Call InitMainTreeview("", m_Commissions)
      Call EnableForm(Me, True)
   End If

   Unload frmAddEditCommissionChart
   Set frmAddEditCommissionChart = Nothing
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim OKClick As Boolean
Dim ID As Long
Dim TableName As String

   
   If m_HasModify Or (m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID") <= 0) Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If

   If trvMain.SelectedItem Is Nothing Then
      glbErrorLog.LocalErrorMsg = ""
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If trvMain.Nodes.Count <= 0 Then
      Exit Sub
   End If
   
   ID = Val(trvMain.SelectedItem.Tag)
            
   Call EnableForm(Me, False)
   frmAddEditCommissionChart.HeaderText = "แก้ไขข้อมูลแผนภูมิการคิดคอมมิตชั่น"
   frmAddEditCommissionChart.ShowMode = SHOW_EDIT
   frmAddEditCommissionChart.ID = ID
   frmAddEditCommissionChart.FK_ID = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID")
   Load frmAddEditCommissionChart
   frmAddEditCommissionChart.Show 1

   If frmAddEditCommissionChart.OKClick Then
      Call EnableForm(Me, False)
      Call LoadCommissionChart(Nothing, m_Commissions, m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID"))
      Call InitMainTreeview("", m_Commissions)
      Call EnableForm(Me, True)
   End If

   Unload frmAddEditCommissionChart
   Set frmAddEditCommissionChart = Nothing

   Call EnableForm(Me, True)
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Function GetIconNo(O As CCommissionChart) As Long
   If O.GetFieldValue("CHILD_COUNT") = 0 Then
      GetIconNo = 2
   Else
      GetIconNo = 1
   End If
End Function
Private Sub GenerateTree(TempColl As Collection, N As Node, NodeID As String, PID As Long, Level As Long)
Dim O As CCommissionChart
Dim Node As Node
Dim NewNodeID As String
Dim L As Long

   For Each O In TempColl
      If O.GetFieldValue("PARENT_ID") = PID Then
         If Level = 0 Then
            Set Node = trvMain.Nodes.add(, tvwFirst, NodeID & O.GetFieldValue("COMMISSION_CHART_ID"), O.GetFieldValue("EMP_NAME") & " " & O.GetFieldValue("EMP_LNAME") & " " & O.GetFieldValue("SALE_NAME") & " " & O.GetFieldValue("RETURN_NAME"), GetIconNo(O))
            Node.Tag = O.GetFieldValue("COMMISSION_CHART_ID")
            Call GenerateTree(TempColl, Node, NodeID & O.GetFieldValue("COMMISSION_CHART_ID"), O.GetFieldValue("COMMISSION_CHART_ID"), Level + 1)
            Call O.SetFieldValue("CHILD_COUNT", Level)
         Else
            NewNodeID = NodeID & "-" & O.GetFieldValue("COMMISSION_CHART_ID")
            Set Node = trvMain.Nodes.add(N, tvwChild, NewNodeID, O.GetFieldValue("EMP_NAME") & "  " & O.GetFieldValue("EMP_LNAME") & " " & O.GetFieldValue("SALE_NAME") & " " & O.GetFieldValue("RETURN_NAME"), GetIconNo(O))
            Node.Tag = O.GetFieldValue("COMMISSION_CHART_ID")
            Call GenerateTree(TempColl, Node, NewNodeID, O.GetFieldValue("COMMISSION_CHART_ID"), Level + 1)
            Call O.SetFieldValue("CHILD_COUNT", Level)
         End If
         Node.Expanded = True
      End If
   Next O
End Sub

Private Sub InitMainTreeview(Caption As String, TempColl As Collection)
   If TempColl Is Nothing Then
      Exit Sub
   End If
   
   ClearTreeView (trvMain.hwnd)
   Call GenerateTree(TempColl, Nothing, "ROOT", -1, 0)
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Me.KeyPreview = True
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_FORM_COLOR
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblMasterFromToNo, MapText("หมายเลข"))
   Call InitNormalLabel(lblMasterFromToDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
   Call txtMasterFromToNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call EnableForm(Me, True)
   
   HasActivate = False
   Set m_Rs = New ADODB.Recordset
   Set m_MasterFromTo = New CMasterFromTo
   Set m_Commissions = New Collection
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

   HasActivate = False
   Me.Caption = HeaderText
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"
   Call InitFormLayout
   
   Exit Sub
   
ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_MasterFromTo = Nothing
   Set m_Commissions = Nothing
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_ID", ID)
      If Not glbDaily.QueryMasterFromTo(m_MasterFromTo, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_MasterFromTo.PopulateFromRS(1, m_Rs)
      
      txtMasterFromToNo.Text = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_NO")
      txtMasterFromToDesc.Text = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_DESC")
      uctlFromDate.ShowDate = m_MasterFromTo.GetFieldValue("VALID_FROM")
      uctlToDate.ShowDate = m_MasterFromTo.GetFieldValue("VALID_TO")
   
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

   If Not VerifyTextControl(lblMasterFromToNo, txtMasterFromToNo, False) Then
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
   
   m_MasterFromTo.ShowMode = ShowMode
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_ID", ID)
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_NO", txtMasterFromToNo.Text)
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_DESC", txtMasterFromToDesc.Text)
   Call m_MasterFromTo.SetFieldValue("VALID_FROM", uctlFromDate.ShowDate)
   Call m_MasterFromTo.SetFieldValue("VALID_TO", uctlToDate.ShowDate)
   Call m_MasterFromTo.SetFieldValue("MASTER_FROMTO_TYPE", DocumentType)
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditMasterFromTo(m_MasterFromTo, IsOK, True, glbErrorLog) Then
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
      ID = m_MasterFromTo.GetFieldValue("MASTER_FROMTO_ID")
      m_MasterFromTo.QueryFlag = 1
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

Private Sub txtMasterFromToDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtMasterFromToNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight - pnlHeader.Height
   SSFrame1.Top = pnlHeader.Height
   pnlHeader.Width = ScaleWidth
   trvMain.Width = ScaleWidth - 2 * trvMain.Left
   trvMain.Height = SSFrame1.Height - trvMain.Top - 620
   cmdAdd.Top = SSFrame1.Height - 580
   cmdEdit.Top = SSFrame1.Height - 580
   cmdDelete.Top = SSFrame1.Height - 580
   cmdOK.Top = SSFrame1.Height - 580
   cmdExit.Top = SSFrame1.Height - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

