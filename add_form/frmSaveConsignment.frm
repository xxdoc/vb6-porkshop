VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmSaveConsignment 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaveConsignment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Rebate.uctlTextBox txtNote 
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   873
      End
      Begin Rebate.uctlTextBox txtDocumentNo 
         Height          =   495
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin Rebate.uctlTextLookup uctlLocationLookup 
         Height          =   435
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin Rebate.uctlTextLookup uctlToLocationLookup 
         Height          =   435
         Left            =   2160
         TabIndex        =   3
         Top             =   2040
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblDocumentNo 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4440
         TabIndex        =   5
         Top             =   2760
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2520
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmSaveConsignment.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   4290
         TabIndex        =   10
         Top             =   2130
         Width           =   1005
      End
      Begin VB.Label lblToLocation 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   4305
         TabIndex        =   8
         Top             =   2580
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmSaveConsignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE

Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public COMMIT_FLAG As String

Public InventorySubType As Long
Public TempCollection As Collection
Public DocumentDate As Date

Private m_Locations As Collection
Private m_ToLocations As Collection
Private m_InventoryDoc As CInventoryDoc
Public ParentForm As Object
Public DocumentType As INVENTORY_DOCTYPE

Private m_Cd As Collection
Private DocAdd As Long
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub
Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblDocumentNo, MapText("เลขที่เอกสาร"))
   Call InitNormalLabel(lblLocation, MapText("จากสถานที่จัดเก็บ"))
   Call InitNormalLabel(lblToLocation, MapText("ไปสถานที่จัดเก็บ"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDocumentNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub
Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyCombo(lblLocation, uctlLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblToLocation, uctlToLocationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDocumentNo, txtDocumentNo, False) Then
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
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_DATE", DocumentDate)
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_NO", txtDocumentNo.Text)
   Call m_InventoryDoc.SetFieldValue("DELIVERY_FEE", 0)
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_TYPE", DocumentType)
   Call m_InventoryDoc.SetFieldValue("COMMIT_FLAG", Check2Flag(0))
   Call m_InventoryDoc.SetFieldValue("EXCEPTION_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("SALE_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("ADJUST_FLAG", "N")
   Call m_InventoryDoc.SetFieldValue("DOCUMENT_DESC", txtNote.Text)
   Call m_InventoryDoc.SetFieldValue("INVENTORY_SUB_TYPE", 1684)
   Call m_InventoryDoc.SetFieldValue("CANCEL_FLAG", Check2Flag(0))
   Call m_InventoryDoc.SetFieldValue("TEMP_SALE_FLAG", Check2Flag(1))
   
   Call EnableForm(Me, False)
   If DocumentType = TRANSFER_DOCTYPE Then
      Call CreateImportExportItems
      Call PopulateGuiID(m_InventoryDoc)
   End If
   
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

   frmAddEditBillingDoc.ConsignmentID = m_InventoryDoc.GetFieldValue("INVENTORY_DOC_ID")
   frmAddEditBillingDoc.ConsignmentNo = txtDocumentNo.Text
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(uctlLocationLookup.MyCombo, m_Locations, , , MASTER_LOCATION)
      Set uctlLocationLookup.MyCollection = m_Locations
      
      Call LoadMaster(uctlToLocationLookup.MyCombo, m_ToLocations, , , MASTER_LOCATION)
      Set uctlToLocationLookup.MyCollection = m_ToLocations
      
      Call LoadConfigDoc(Nothing, m_Cd)

      Call cmdAuto_Click
            
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   End If
End Sub
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
'   Set TempCollection = New Collection
   Set m_Locations = New Collection
   Set m_ToLocations = New Collection
   Set m_InventoryDoc = New CInventoryDoc
   Set m_Cd = New Collection
End Sub
Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
'   Set TempCollection = Nothing
   Set m_Locations = Nothing
   Set m_ToLocations = Nothing
   Set m_InventoryDoc = Nothing
   Set m_Cd = Nothing
End Sub
Private Sub txtNote_Change()
   m_HasModify = True
End Sub
Private Sub uctlToLocationLookup_Change()
   m_HasModify = True
End Sub
Private Sub uctlLocationLookup_Change()
   m_HasModify = True
End Sub
Private Sub CreateImportExportItems()
Dim Ti As CTransferItem
Dim EnpAddress As CTransferItem
Dim Di As CDocItem
Dim Ei As CLotItem
Dim II As CLotItem

   Set m_InventoryDoc.ImportExportItems = Nothing
   Set m_InventoryDoc.ImportExportItems = New Collection
   
   For Each Di In TempCollection
         Set Ei = New CLotItem
         Set II = New CLotItem
         Set EnpAddress = New CTransferItem
         
         Ei.Flag = "A"
         II.Flag = "A"
         EnpAddress.Flag = "A"
         Set EnpAddress.ExportItem = Ei
         Set EnpAddress.ImportItem = II

         Call m_InventoryDoc.TransferItems.add(EnpAddress)

         EnpAddress.ExportItem.PART_TYPE = Di.GetFieldValue("STOCK_TYPE")
         EnpAddress.ExportItem.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
         EnpAddress.ExportItem.LOCATION_ID = uctlLocationLookup.MyCombo.ItemData(Minus2Zero(uctlLocationLookup.MyCombo.ListIndex))
         
         EnpAddress.ExportItem.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
         EnpAddress.ExportItem.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
         EnpAddress.ExportItem.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("TOTAL_INCLUDE_PRICE")
         
         EnpAddress.ExportItem.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
         EnpAddress.ExportItem.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
         EnpAddress.ExportItem.UNIT_TRAN_NAME = Di.GetFieldValue("UNIT_TRAN_NAME")
         
         EnpAddress.ExportItem.PART_TYPE_NAME = Di.GetFieldValue("STOCK_DESC")
         EnpAddress.ExportItem.LOCATION_NAME = uctlLocationLookup.MyCombo.Text
         EnpAddress.ExportItem.PART_NO = Di.GetFieldValue("STOCK_NO")
         EnpAddress.ExportItem.PART_DESC = Di.GetFieldValue("STOCK_DESC")
         EnpAddress.ExportItem.TX_TYPE = "E"
         EnpAddress.ExportItem.MULTIPLIER = -1
         
         EnpAddress.ImportItem.PART_TYPE = Di.GetFieldValue("STOCK_TYPE")
         EnpAddress.ImportItem.PART_ITEM_ID = Di.GetFieldValue("PART_ITEM_ID")
         EnpAddress.ImportItem.LOCATION_ID = uctlToLocationLookup.MyCombo.ItemData(Minus2Zero(uctlToLocationLookup.MyCombo.ListIndex))

         EnpAddress.ImportItem.TX_AMOUNT = Di.GetFieldValue("ITEM_AMOUNT")
         EnpAddress.ImportItem.AVG_PRICE = Di.GetFieldValue("AVG_PRICE")
         EnpAddress.ImportItem.TOTAL_INCLUDE_PRICE = Di.GetFieldValue("TOTAL_INCLUDE_PRICE")

         EnpAddress.ImportItem.PART_TYPE_NAME = Di.GetFieldValue("STOCK_DESC")
         EnpAddress.ImportItem.LOCATION_NAME = uctlToLocationLookup.MyCombo.Text
         EnpAddress.ImportItem.PART_NO = Di.GetFieldValue("STOCK_NO")
         EnpAddress.ImportItem.PART_DESC = Di.GetFieldValue("STOCK_DESC")
         EnpAddress.ImportItem.TX_TYPE = "I"
         EnpAddress.ImportItem.MULTIPLIER = 1
         
         EnpAddress.ImportItem.UNIT_TRAN_ID = Di.GetFieldValue("UNIT_TRAN_ID")
         EnpAddress.ImportItem.UNIT_MULTIPLE = Di.GetFieldValue("UNIT_MULTIPLE")
         EnpAddress.ImportItem.UNIT_TRAN_NAME = Di.GetFieldValue("UNIT_TRAN_NAME")
         
         Set EnpAddress = Nothing
   Next Di
   
   For Each Ti In m_InventoryDoc.TransferItems
      Set Ei = Ti.ExportItem
      Set II = Ti.ImportItem
      
      Ei.Flag = Ti.Flag
      II.Flag = Ti.Flag
      
      Call m_InventoryDoc.ImportExportItems.add(Ei)
      Call m_InventoryDoc.ImportExportItems.add(II)
   Next Ti
End Sub
Private Sub PopulateGuiID(BD As CInventoryDoc)
Dim Di As CLotItem
Dim I As Long
Dim TempID As Long
   I = 0
   For Each Di In BD.ImportExportItems
      If Di.Flag = "A" Then
         I = I + 1
         If (I Mod 2) = 1 Then
            Di.LINK_ID = GetNextGuiID(BD)
            TempID = Di.LINK_ID
         Else
            Di.LINK_ID = TempID
         End If
         
      End If
   Next Di
End Sub
Private Function GetNextGuiID(BD As CInventoryDoc) As Long
Dim Di As CLotItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.ImportExportItems
      If Di.LINK_ID > MaxId Then
         MaxId = Di.LINK_ID
      End If
   Next Di

   GetNextGuiID = MaxId + 1
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
   
   DocumentType = TRANSFER_DOCTYPE
   ID = ConvertDocToConfigNo(2, DocumentType, -1)
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
