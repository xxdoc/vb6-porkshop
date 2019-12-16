VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmConfigDoc 
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6660
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3240
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5715
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboDocumentType 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1200
         Width           =   3375
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtDigitAmount 
         Height          =   405
         Left            =   3720
         TabIndex        =   3
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtPreFix 
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   2040
         Width           =   2415
         _ExtentX        =   1931
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtRunningNo 
         Height          =   405
         Left            =   4440
         TabIndex        =   4
         Top             =   2040
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
      End
      Begin PorkShop.uctlTextBox txtLastNo 
         Height          =   405
         Left            =   3840
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
      End
      Begin VB.Label lblLastNo 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   840
         Width           =   2505
      End
      Begin VB.Label lblRunningNo 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   4440
         TabIndex        =   12
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblPreFix 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblDigitAmount 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   1680
         Width           =   585
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3450
         TabIndex        =   7
         Top             =   2580
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1800
         TabIndex        =   6
         Top             =   2580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblDocumentType 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmConfigDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Cd As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public AllDocType As Collection
Private Sub cboDocumentType_Click()
Dim ID As Long
Dim Cd As CConfigDoc
   
   ID = cboDocumentType.ItemData(Minus2Zero(cboDocumentType.ListIndex))
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         txtLastNo.Text = Cd.GetFieldValue("LAST_NO")
         txtPreFix.Text = Cd.GetFieldValue("PREFIX")
         txtDigitAmount.Text = Cd.GetFieldValue("DIGIT_AMOUNT")
         txtRunningNo.Text = Cd.GetFieldValue("RUNNING_NO")
         
      Else
         txtLastNo.Text = ""
         txtPreFix.Text = ""
         txtDigitAmount.Text = ""
         txtRunningNo.Text = ""
      End If
   End If
   
   m_HasModify = False
End Sub
Private Sub cboDocumentType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents

      Call EnableForm(Me, False)
      
      Call LoadConfigDoc(Nothing, m_Cd)
      Call GenerateAllConfigDoc
      Call LoadDocType
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub
Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Cd = New Collection
   Set AllDocType = New Collection
   
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   Call InitNormalLabel(lblDocumentType, MapText("ประเภทเอกสาร"))
   Call InitNormalLabel(lblLastNo, MapText("หมายเลขสุดท้าย"))
   Call InitNormalLabel(lblPreFix, MapText("Prefix"))
   Call InitNormalLabel(lblDigitAmount, MapText("หลัก"))
   Call InitNormalLabel(lblRunningNo, MapText("RunNo"))
   
   Call txtDigitAmount.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   Call txtRunningNo.SetTextLenType(TEXT_INTEGER, glbSetting.AMOUNT_LEN)
   txtLastNo.Enabled = False
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   Call InitCombo(cboDocumentType)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))

End Sub
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If

   OKClick = False
   Unload Me
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim I As Long
Dim ID As Long
Dim Cd As CConfigDoc
   
   If Not VerifyCombo(lblDocumentType, cboDocumentType, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblDigitAmount, txtDigitAmount, False) Then
      Exit Function
   End If
   
   If Not VerifyTextControl(lblRunningNo, txtRunningNo, True) Then
      Exit Function
   End If
   
   ID = cboDocumentType.ItemData(Minus2Zero(cboDocumentType.ListIndex))
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Cd Is Nothing Then
         Set Cd = New CConfigDoc
         Cd.Flag = "A"
         Call Cd.SetFieldValue("CONFIG_DOC_TYPE", ID)
      Else
         Cd.Flag = "E"
      End If
   End If
      
      
   If Cd.Flag = "A" Then
      Cd.ShowMode = SHOW_ADD
   ElseIf Cd.Flag = "E" Then
      Cd.ShowMode = SHOW_EDIT
   End If
   Call Cd.SetFieldValue("PREFIX", txtPreFix.Text)
   Call Cd.SetFieldValue("DIGIT_AMOUNT", txtDigitAmount.Text)
   Call Cd.SetFieldValue("RUNNING_NO", txtRunningNo.Text)
   
   Call EnableForm(Me, False)

   Call Cd.AddEditData
  
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub cmdOK_Click()
   If cmdOK.Enabled = False Then
      Exit Sub
   End If
   Call SaveData
   
   m_HasModify = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub GenerateAllConfigDoc()
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
Dim MenuMask As String
   
   MenuMask = "YYYYYY"
   '1
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("ใบ QUATATION (ขาย)")
'   D.KEY_ID = 1
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.add(D)
'   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบ PO (ขาย)")
   D.KEY_ID = 2
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' DO load From DataBase
   Dim Mr As CMasterRef
   Dim TempColl  As Collection
   Set TempColl = New Collection
   Call LoadMaster(Nothing, TempColl, , , MASTER_INVOICE_SUB)
   For Each Mr In TempColl
      Set D = New CMenuItem
      D.MENU_TEXT = Mr.KEY_NAME
      D.KEY_ID = 100 + Mr.KEY_ID
      D.PARENT_KEY = ""
      D.ICON_INDEX1 = 1
      D.ICON_INDEX2 = 2
      Call AllDocType.add(D)
      Set D = Nothing
   Next Mr
   Set Mr = Nothing
   Set TempColl = Nothing
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' DO load From DataBase
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบเสร็จ (ขายสด)")
   D.KEY_ID = 3
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบเสร็จ (แนบบิลใบส่งของ)")
   D.KEY_ID = 9
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบเสร็จ (รับชำระ)")
   D.KEY_ID = 4
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("เพิ่มหนี้ (ขาย)")
   D.KEY_ID = 5
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ลดหนี้ (ขาย)")
   D.KEY_ID = 6
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ลดหนี้ รับคืนสินค้า (ขาย)")
   D.KEY_ID = 7
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบวางบิล (ขาย)")
   D.KEY_ID = 8
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   '19
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("ใบ QUATATION (ซื้อ)")
'   D.KEY_ID = 19
'   D.PARENT_KEY = ""
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call AllDocType.add(D)
'   Set D = Nothing
      
      
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบ PO (ซื้อ)")
   D.KEY_ID = 20
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบรับสินค้า (ซื้อ)")
   D.KEY_ID = 21
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบเสร็จ (ซื้อสด)")
   D.KEY_ID = 22
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบเสร็จ (รับชำระซื้อ)")
   D.KEY_ID = 23
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("เพิ่มหนี้ (ซื้อ)")
   D.KEY_ID = 24
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ลดหนี้ (ซื้อ)")
   D.KEY_ID = 25
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ลดหนี้ ส่งคืนสินค้า (ซื้อ)")
   D.KEY_ID = 26
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบวางบิล (ซื้อ)")
   D.KEY_ID = 27
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบนำเข้า")
   D.KEY_ID = 50
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบเบิก")
   D.KEY_ID = 51
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   '===
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบโอน")
   D.KEY_ID = 52
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบปรับยอด")
   D.KEY_ID = 53
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ใบผลิต")
   D.KEY_ID = 1000
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call AllDocType.add(D)
   Set D = Nothing
   
   '====
   TempCount = AllDocType.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call AllDocType.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Cd = Nothing
   Set AllDocType = Nothing
   If m_Rs.State = adStateOpen Then
      Call m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub LoadDocType()
Dim Mu As CMenuItem
Dim I As Long
   I = 0
   cboDocumentType.Clear
   cboDocumentType.AddItem ("")
   
   For Each Mu In AllDocType
      I = I + 1
      cboDocumentType.AddItem (Mu.MENU_TEXT)
      cboDocumentType.ItemData(I) = Mu.KEY_ID
   Next
End Sub
Private Sub txtDigitAmount_Change()
   m_HasModify = True
End Sub
Private Sub txtLastNo_Change()
   m_HasModify = True
End Sub
Private Sub txtPreFix_Change()
   m_HasModify = True
End Sub
Private Sub txtRunningNo_Change()
   m_HasModify = True
End Sub
