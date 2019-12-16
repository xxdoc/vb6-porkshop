VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCopyPoToDo 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmCopyPoToDo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3405
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6006
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin PorkShop.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   900
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   2280
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin PorkShop.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1980
         TabIndex        =   3
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin PorkShop.uctlDate uctlToDate 
         Height          =   405
         Left            =   7170
         TabIndex        =   1
         Top             =   900
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   5880
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   7800
         TabIndex        =   4
         Top             =   2700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCopyPoToDo.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3720
         TabIndex        =   11
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   360
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   9
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   9495
         TabIndex        =   5
         Top             =   2700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCopyPoToDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HeaderText As String
Private DistinctBillingDocIDColl As Collection
Private HaveBillingDocIDColl As Collection
Private m_Cd As Collection
Private m_Sc As CStockCode
Private m_Products As Collection
Private Sub cmdOK_Click()
   Unload Me
End Sub
Private Sub cmdStart_Click()
Dim Status As Boolean
Dim PartItemID As Long
Dim DriverID As Long
Dim TranSportorID As Long
   
   If Not VerifyLockDate(uctlFromDate.ShowDate, uctlToDate.ShowDate) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   If Not VerifyLockInvoiceDate(uctlFromDate.ShowDate, uctlToDate.ShowDate) Then
      glbErrorLog.LocalErrorMsg = MapText("ไม่สามารถเปลี่ยนแปลงเอกสารตามวันที่เอกสารที่เลือกได้ กรุณาติดต่อผู้ดูแลระบบ หรือผู้มีสิทธิ์กำหนดวันที่เอกสารได้")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call LoadDisTinctPOID(HaveBillingDocIDColl, uctlFromDate.ShowDate, uctlToDate.ShowDate, "(" & INVOICE_DOCTYPE & "," & RECEIPT1_DOCTYPE & ")")
   
   Call LoadDisTinctBillingDocID(DistinctBillingDocIDColl, , , , PO_DOCTYPE, uctlFromDate.ShowDate, uctlToDate.ShowDate, 1)
   
   Call LoadStockCode(Nothing, m_Products)
   
   
   Call glbDaily.StartTransaction
      
   Me.Enabled = False
   
   '-------------------------------------------------------------------------------------------------------------------------------
   
   Status = GenerateDoFromPo
   
   '-------------------------------------------------------------------------------------------------------------------------------
   
   
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
   
   Call cmdOK_Click
   Exit Sub
   
End Sub

Private Sub Form_Activate()
   
   Me.Refresh
   DoEvents
   
   uctlFromDate.ShowDate = Now
   uctlToDate.ShowDate = Now
   
   Call LoadConfigDoc(Nothing, m_Cd)
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

Private Sub ResetStatus()
   prgProgress.Max = 100
   prgProgress.Min = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = "คัดลอกของมูลจากเอกสารใบสั่งซื้อมาเป็นบิลขาย ตามวันที่ส่งของ"
   pnlHeader.Caption = Me.Caption
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   Call InitNormalLabel(lblFromDate, "จากวันที่", RGB(255, 0, 0))
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblToDate, "ถึงวันที่", RGB(255, 0, 0))
   
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
  ' cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
  ' Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   Call EnableForm(Me, False)
   
   Set DistinctBillingDocIDColl = New Collection
   Set HaveBillingDocIDColl = New Collection
   Set m_Cd = New Collection
   Set m_Products = New Collection
   
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DistinctBillingDocIDColl = Nothing
   Set HaveBillingDocIDColl = Nothing
   Set m_Cd = Nothing
   Set m_Products = Nothing
End Sub
Private Function GenerateDoFromPo() As Boolean
Dim PoID As CBillingDoc
Dim m_BillingDoc As CBillingDoc
Dim m_Rs As ADODB.Recordset
Dim ItemCount As Long
Dim IsOK As Boolean
Dim AparMas As CAPARMas
Dim PrevKey1 As String
Dim ConFigDocType As Long
Dim RunningNo As Long
Dim HeadNo As String
Dim DocItem As CDocItem
Dim Ivd As CInventoryDoc
Dim Cd As CConfigDoc
Dim FormatAmount As String
Dim I As Long
Dim TempBd As CBillingDoc
Dim HaveMoreToUse As Boolean

   GenerateDoFromPo = False
   
   If DistinctBillingDocIDColl.Count > 0 Then
      prgProgress.Max = DistinctBillingDocIDColl.Count
   End If
   
   I = 0
   For Each PoID In DistinctBillingDocIDColl
      I = I + 1
      prgProgress.Value = I
      txtPercent.Text = MyDiff(I, DistinctBillingDocIDColl.Count) * 100
      Me.Refresh
      Set TempBd = GetObject("CBillingDoc", HaveBillingDocIDColl, Trim(Str(PoID.BILLING_DOC_ID)), False)
      If (TempBd Is Nothing) Then
         Set m_BillingDoc = New CBillingDoc
         Set m_Rs = New ADODB.Recordset
         m_BillingDoc.BILLING_DOC_ID = PoID.BILLING_DOC_ID
         m_BillingDoc.QueryFlag = 1
         If Not glbDaily.QueryBillingDocOnlyDo(m_BillingDoc, m_Rs, ItemCount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Function
         End If
         
         Call m_BillingDoc.PopulateFromRS(1, m_Rs)
         
         HaveMoreToUse = True
         For Each DocItem In m_BillingDoc.DocItems
            If DocItem.GetFieldValue("DOC_ITEM_TYPE") = 1 Then
                     
'               If Not (LoadCheckBalance(DocItem.GetFieldValue("ITEM_AMOUNT"), DocItem.GetFieldValue("LOCATION_ID"), DocItem.GetFieldValue("PART_ITEM_ID"), DocItem.GetFieldValue("STOCK_NO") & " - คลัง " & DocItem.GetFieldValue("LOCATION_NAME"))) Then
'                  HaveMoreToUse = False
'               End If
               
               DocItem.Flag = "A"
               Call DocItem.SetFieldValue("PO_ID", m_BillingDoc.BILLING_DOC_ID)
               Call DocItem.SetFieldValue("PO_NO", m_BillingDoc.DOCUMENT_NO)
               
               Set m_Sc = GetObject("CStockCode", m_Products, Trim(Str(DocItem.GetFieldValue("PART_ITEM_ID"))))
               
                'If m_Sc.CHK_STD_COST = "Y" Then        'ถ้าเป็น Standard แล้วให้นำต้นทุน Standard เป็นต้นทุนขายด้วยทันที
                '  Call DocItem.SetFieldValue("CAPITAL_AMOUNT", m_Sc.COST_PER_AMOUNT)
                  'Call DocItem.SetFieldValue("TOTAL_INCLUDE_PRICE", m_Sc.COST_PER_AMOUNT * DocItem.GetFieldValue("ITEM_AMOUNT"))
               'End If
               
            Else
               DocItem.Flag = "I"
            End If
         Next DocItem
         
         If HaveMoreToUse Then
            m_BillingDoc.ShowMode = SHOW_ADD
            m_BillingDoc.BILLING_DOC_ID = -1
            m_BillingDoc.DOCUMENT_DATE = m_BillingDoc.Due_Date
            
            Set AparMas = m_CustomerColl(Trim(Str(m_BillingDoc.APAR_MAS_ID)))
            m_BillingDoc.CREDIT = AparMas.CREDIT
            m_BillingDoc.Due_Date = DateAdd("D", m_BillingDoc.CREDIT, m_BillingDoc.DOCUMENT_DATE)
            '------------------------------------------------------------------------------------------------------------------------------------------------
            If m_BillingDoc.DOCUMENT_SUB_TYPE > 0 Then
               'ขายเชื่อ   m_BillingDoc.DOCUMENT_SUB_TYPE
               m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE
               
            Else
               'ขายสด
               m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE
            End If
            'm_BillingDoc.DOCUMENT_NO
            '------------------------------------------------------------------------------------------------------------------------------------------------
            If PrevKey1 <> Trim(Str(PoID.DOCUMENT_SUB_TYPE)) Then
               '------------------------------------ Update ของเดิมก่อนเปล่าประเภทเอกสาร
               If RunningNo > 0 Then
                  Set Cd = New CConfigDoc
                  Call Cd.SetFieldValue("RUNNING_NO", RunningNo)
                  Call Cd.SetFieldValue("LAST_NO", HeadNo & Format(Trim(Str(RunningNo)), FormatAmount))
                  Call Cd.SetFieldValue("CONFIG_DOC_TYPE", ConFigDocType)
                  Call Cd.UpdateRunningNo
               End If
               '------------------------------------ Update ของเดิมก่อนเปล่าประเภทเอกสาร
               
               m_BillingDoc.DOCUMENT_NO = GetDocumentNo(m_BillingDoc.DOCUMENT_TYPE, m_BillingDoc.DOCUMENT_SUB_TYPE, m_BillingDoc.DOCUMENT_DATE, HeadNo, RunningNo, ConFigDocType, FormatAmount)
               m_BillingDoc.RUNNING_NO = RunningNo
            Else
               RunningNo = RunningNo + 1
               m_BillingDoc.DOCUMENT_NO = HeadNo & Format(Trim(Str(RunningNo)), FormatAmount)
            End If
            PrevKey1 = Trim(Str(PoID.DOCUMENT_SUB_TYPE))
            '------------------------------------------------------------------------------------------------------------------------------------------------
            
            'm_BillingDoc.APAR_MAS_ID
            'm_BillingDoc.DEPARTMENT_ID
            'm_BillingDoc.BILLING_ADDRESS_ID
            'm_BillingDoc.ENTERPRISE_ADDRESS_ID
            'm_BillingDoc.DISCOUNT_AMOUNT
            'm_BillingDoc.EXT_DISCOUNT_AMOUNT
            'm_BillingDoc.EXT_DISCOUNT_PERCENT
            'm_BillingDoc.ADDITION_AMOUNT
            'm_BillingDoc.TOTAL_AMOUNT
            'm_BillingDoc.TOTAL_PRICE
            
            'm_BillingDoc.VAT_PERCENT
            'm_BillingDoc.VAT_AMOUNT
            
            'm_BillingDoc.WH_PERCENT
            'm_BillingDoc.WH_AMOUNT
            'm_BillingDoc.PAY_AMOUNT
            'm_BillingDoc.PAID_AMOUNT
            
            'm_BillingDoc.CREDIT_AMOUNT
            'm_BillingDoc.DEBIT_AMOUNT
            'm_BillingDoc.FEE_AMOUNT
            
            'm_BillingDoc.NOTE
            'm_BillingDoc.REFER_TEXT
            'm_BillingDoc.REFER_DESC
            ''m_BillingDoc.SALE_BY
            'm_BillingDoc.CUS_PO
            'm_BillingDoc.BRANCH_ADDRESS
            'm_BillingDoc.DOCUMENT_RETURN
            
            'm_BillingDoc.DRIVER_ID
            'm_BillingDoc.CAR_LICENSE_ID
            'm_BillingDoc.TRANSPORTOR_ID
            
            Call PopulateGuiID(m_BillingDoc)
            
            If m_BillingDoc.DOCUMENT_TYPE = INVOICE_DOCTYPE Then 'ใบส่งสินค้าขาย
               Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, 1, 10)
            ElseIf m_BillingDoc.DOCUMENT_TYPE = RECEIPT1_DOCTYPE Then   'ใบเสร็จขายสด
               Call glbDaily.DO2InventoryDoc(m_BillingDoc, Ivd, 1, 21)
            End If
            
            If Not glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog) Then
               glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
               Call glbDaily.RollbackTransaction
               Exit Function
            End If
            m_BillingDoc.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
            
            If Not glbDaily.AddEditBillingDoc(m_BillingDoc, IsOK, False, glbErrorLog) Then
               glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
               Call glbDaily.RollbackTransaction
               Exit Function
            End If
         End If
      End If
   Next PoID
   
   If RunningNo > 0 Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", RunningNo)
      Call Cd.SetFieldValue("LAST_NO", HeadNo & Format(Trim(Str(RunningNo)), FormatAmount))
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", ConFigDocType)
      Call Cd.UpdateRunningNo
   End If
            
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   GenerateDoFromPo = True
End Function
Private Function GetDocumentNo(DocumentType As Long, DocumentSubType As Long, DocumentDate As Date, HeadNo As String, RunningNo As Long, ConFigDocType As Long, TempStr As String) As String
Dim ID As Long
Dim Cd As CConfigDoc
Dim I As Long
   
   GetDocumentNo = ""
   
   ID = ConvertDocToConfigNo(1, DocumentType, DocumentSubType)
   If ID <= 0 Then
      glbErrorLog.LocalErrorMsg = "ไม่สามารถดำเนินการต่อได้ เนื่องจากระบบจำเป็นที่จะต้องตั้งหมายเลขเอกสารอัตโนมัติไว้ก่อน"
      glbErrorLog.ShowUserError
      Exit Function
   End If
   If ID > 0 Then
      Set Cd = GetObject("CConfigDoc", m_Cd, Trim(Str(ID)), False)
      If Not (Cd Is Nothing) Then
         Dim TempCd As CConfigDoc
         ''''''''''''''
         
         GetDocumentNo = Cd.GetFieldValue("PREFIX")
         TempStr = ""
         For I = 1 To Cd.GetFieldValue("DIGIT_AMOUNT")
            TempStr = TempStr & "0"
         Next I
         
         HeadNo = GetDocumentNo
         GetDocumentNo = GetDocumentNo & Format(Cd.GetFieldValue("RUNNING_NO") + 1, TempStr)
         RunningNo = Cd.GetFieldValue("RUNNING_NO") + 1
         ConFigDocType = ID
      ElseIf Cd Is Nothing Then
         glbErrorLog.LocalErrorMsg = "ไม่สามารถดำเนินการต่อได้ เนื่องจากระบบจำเป็นที่จะต้องตั้งหมายเลขเอกสารอัตโนมัติไว้ก่อน"
         glbErrorLog.ShowUserError
         Exit Function
      End If
   End If
End Function
Private Sub PopulateGuiID(BD As CBillingDoc)
Dim Di As CDocItem

   For Each Di In BD.DocItems
      If Di.Flag = "A" Then
         Call Di.SetFieldValue("LINK_ID", GetNextGuiID(BD))
      End If
   Next Di
End Sub
Private Function GetNextGuiID(BD As CBillingDoc) As Long
Dim Di As CDocItem
Dim MaxId As Long

   MaxId = 0
   For Each Di In BD.DocItems
      If Di.GetFieldValue("LINK_ID") > MaxId Then
         MaxId = Di.GetFieldValue("LINK_ID")
      End If
   Next Di

   GetNextGuiID = MaxId + 1
End Function

