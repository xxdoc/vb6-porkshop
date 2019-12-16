VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportDoc 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmImportDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3765
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   6641
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboImportType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1020
         Width           =   3105
      End
      Begin Rebate.uctlTextBox txtFileName 
         Height          =   435
         Left            =   1860
         TabIndex        =   11
         Top             =   1470
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   767
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   1920
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Rebate.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   2250
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   9780
         Top             =   750
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   525
         Left            =   6270
         TabIndex        =   14
         Top             =   780
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":27A2
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFileName 
         Height          =   405
         Left            =   8670
         TabIndex        =   12
         Top             =   1470
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   2910
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmImportDoc.frx":2DD6
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   10
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblMasterName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   3
         Top             =   2910
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmImportDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Private m_ExcelApp As Object
Private m_ExcelSheet As Object
Private Sub cmdFileName_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Access Files (*.XLS)|*..xls;*.XLS;"
   dlgAdd.DialogTitle = "Select access file to import"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   txtFileName.Text = dlgAdd.FileName
   m_HasModify = True
End Sub
Private Sub cmdStart_Click()
Dim TempID As Long

   If Not VerifyCombo(lblFileName, cboImportType) Then
      Exit Sub
   End If
   If Not VerifyTextControl(lblMasterName, txtFileName) Then
      Exit Sub
   End If
   
   TempID = cboImportType.ItemData(Minus2Zero(cboImportType.ListIndex))
   If TempID <= 0 Then
      Exit Sub
   End If
   
   If TempID = 1 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      Call ImportARBalance
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
      
      cmdExit_Click
   ElseIf TempID = 2 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 3 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      'Call ImportCustomer
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   ElseIf TempID = 4 Then
      Call EnableForm(Me, False)
      m_ExcelApp.Workbooks.Close
      m_ExcelApp.Workbooks.Open (txtFileName.Text)
      
      'Call ImportAdjust
      
      m_ExcelApp.Workbooks.Close
      Call EnableForm(Me, True)
   End If
End Sub
Private Function GetVal(Row As Long, Col As Long) As Double
On Error Resume Next
   GetVal = m_ExcelSheet.Cells(Row, Col).Value
End Function

Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitImportType(cboImportType)
      
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
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      'Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
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
   pnlHeader.Caption = "อิมพอร์ตข้อมูล"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "ประเภท")
   Call InitNormalLabel(lblMasterName, "ชื่อไฟล์")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   Call txtFileName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtFileName.Enabled = False
   
   Call InitCombo(cboImportType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdFileName.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   Call InitMainButton(cmdFileName, MapText("..."))
   
   Call ResetStatus
End Sub
Private Sub cmdExit_Click()
   
   OKClick = False
   m_ExcelApp.Workbooks.Close
   Unload Me
End Sub
Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_Rs = New ADODB.Recordset
   
   Set m_ExcelApp = CreateObject("Excel.application")
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub ImportARBalance()
On Error GoTo ErrorHandler
Dim MaxRow As Long
Dim MaxCol As Long
Dim ID As Long
Dim FieldNames() As String
Dim FieldTypes() As String
Dim I As Long
Dim TabField As String
Dim StateMent As String
Dim NewValue As String
Dim Row As Long
Dim Col As Long
Dim ErrorCount As Long
Dim SuccessCount As Long
Dim ProgressCount As Long
Dim ErrorFlag As Boolean
Dim ServerDtm As String
Dim HasBegin As Boolean
Dim IsOK As Boolean
Dim BD As CBillingDoc
Dim Cm As CAPARMas
Dim TempCus As CAPARMas
Dim Di As CDocItem
Dim m_Customers As Collection
Dim TempSetDate As Date
Dim TempID As Long
Dim CREDIT As Long

   HasBegin = False
   
   Set m_Customers = New Collection
   
   For Each Cm In m_CustomerColl
      Set TempCus = New CAPARMas
      TempCus.APAR_MAS_ID = Cm.APAR_MAS_ID
      TempCus.APAR_CODE = Cm.APAR_CODE
      TempCus.CREDIT = Cm.CREDIT
      Call m_Customers.add(TempCus, Trim(TempCus.APAR_CODE))
   Next
   ID = 1
   
   Set TempCus = Nothing
   
   Set m_ExcelSheet = m_ExcelApp.Sheets(ID)
      
   MaxRow = m_ExcelSheet.UsedRange.Rows.Count
   MaxCol = m_ExcelSheet.UsedRange.Columns.Count

   ReDim FieldNames(MaxCol)
   ReDim FieldTypes(MaxCol)
   
   Call EnableForm(Me, False)
   cmdStart.Enabled = False
   cmdExit.Enabled = False
   
       
   ProgressCount = 0
   ErrorCount = 0
   SuccessCount = 0
   
   prgProgress.Min = 1
   prgProgress.Max = (MaxRow) + 1
   
   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
      
   For Row = 2 To MaxRow
      DoEvents
      If Len(Trim(m_ExcelSheet.Cells(Row, 2).Value)) = 0 Then
         Exit For
      End If
      Set BD = New CBillingDoc
      
      Set Di = New CDocItem
      
      BD.ShowMode = SHOW_ADD
      
      BD.DOCUMENT_NO = Trim(m_ExcelSheet.Cells(Row, 2).Value)
      TempSetDate = m_ExcelSheet.Cells(Row, 3).Value
      BD.DOCUMENT_DATE = TempSetDate
      BD.TICKET_FLAG = "N"
      BD.CANCEL_FLAG = "N"
      
      If Val(m_ExcelSheet.Cells(Row, 5).Value) > 0 Then
         BD.DOCUMENT_TYPE = INVOICE_DOCTYPE
         BD.TOTAL_PRICE = Val(m_ExcelSheet.Cells(Row, 5).Value)
      Else
         BD.DOCUMENT_TYPE = RETURN_DOCTYPE
         BD.TOTAL_PRICE = -1 * Val(m_ExcelSheet.Cells(Row, 5).Value)
      End If
      
      
      BD.DEPARTMENT_ID = -1
      BD.SALE_BY = -1
      BD.BILLING_ADDRESS_ID = -1
      BD.ENTERPRISE_ADDRESS_ID = -1
      BD.CUSTOMER_BRANCH = -1
      BD.BRANCH_ADDRESS = -1
      
      Set Cm = GetObject("CAPARMas", m_Customers, Trim(m_ExcelSheet.Cells(Row, 1).Value), False)
      
      If Not (Cm Is Nothing) Or Len(Trim(m_ExcelSheet.Cells(Row, 1).Value)) = 0 Then
         If Len(Trim(m_ExcelSheet.Cells(Row, 1).Value)) > 0 Then
            BD.APAR_MAS_ID = Cm.APAR_MAS_ID
            CREDIT = Cm.CREDIT
         Else
            BD.APAR_MAS_ID = TempID
         End If
         
         TempSetDate = -1
         
         If StringToDateCheckError(m_ExcelSheet.Cells(Row, 4).Value) > 0 Then
            TempSetDate = m_ExcelSheet.Cells(Row, 4).Value
            BD.Due_Date = TempSetDate
         Else
            If Cm Is Nothing Then
               BD.Due_Date = DateAdd("D", CREDIT, BD.DOCUMENT_DATE)
            Else
               BD.Due_Date = DateAdd("D", Cm.CREDIT, BD.DOCUMENT_DATE)
            End If
         End If
         
         TempID = BD.APAR_MAS_ID

         
         If TempID > 0 Then
            Di.Flag = "A"
            Call Di.SetFieldValue("STOCK_TYPE", -1)
            Call Di.SetFieldValue("PART_ITEM_ID", -1)
            Call Di.SetFieldValue("LOCATION_ID", -1)
            Call Di.SetFieldValue("ITEM_AMOUNT", 1)
            
            If Val(m_ExcelSheet.Cells(Row, 5).Value) > 0 Then
               Call Di.SetFieldValue("TOTAL_PRICE", Val(m_ExcelSheet.Cells(Row, 5).Value))
            Else
               Call Di.SetFieldValue("TOTAL_PRICE", -1 * Val(m_ExcelSheet.Cells(Row, 5).Value))
            End If
            
            Call Di.SetFieldValue("SELL_TYPE", 3)
            Call Di.SetFieldValue("ITEM_DESC", "รายการยอดยกมา")
            Call Di.SetFieldValue("AVG_PRICE", Di.GetFieldValue("TOTAL_PRICE"))
            
            Call BD.DocItems.add(Di)
            
            Call glbDaily.AddEditBillingDoc(BD, IsOK, False, glbErrorLog)
         End If
      Else
         glbErrorLog.LocalErrorMsg = "ไม่พบ " & Trim(m_ExcelSheet.Cells(Row, 1).Value)
         'glbErrorLog.ShowUserError
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      End If
      
      Set Di = Nothing
      Set BD = Nothing
      
      ProgressCount = ProgressCount + 1
      prgProgress.Value = ProgressCount
      txtPercent.Text = MyDiff(ProgressCount, MaxRow) * 100
      Me.Refresh
   Next Row
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   HasBegin = False
   
   Set m_ExcelSheet = Nothing
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   
   glbErrorLog.LocalErrorMsg = "การตั้งยอดยกมาสำเร็จ"
   glbErrorLog.ShowUserError
   Exit Sub
   
ErrorHandler:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
   End If
   
   Call EnableForm(Me, True)
   
   cmdStart.Enabled = True
   cmdExit.Enabled = True
   
   glbErrorLog.LocalErrorMsg = Err.Description
   glbErrorLog.ShowUserError
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
