VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmCreateBillingDocPack 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmCreateBillingDocPack.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   2685
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   4736
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   840
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
      Begin PorkShop.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   1
         Top             =   1170
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
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1890
         TabIndex        =   2
         Top             =   1830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmCreateBillingDocPack.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   8
         Top             =   1290
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   7
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   3
         Top             =   1830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCreateBillingDocPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public m_BillingDoc As CBillingDoc
Public DocumentDate As Date

Private m_Cd As Collection
Private DocAdd As Long

Private Sub cmdStart_Click()
On Error GoTo Errorhanderror
Dim TempBd As CBillingDoc
Dim TempBDEX As CBillingDoc
Dim Rcp As CRcpCnDn_Item
Dim m_Rs As ADODB.Recordset
Dim I As Long
Dim HasBegin As Boolean
Dim ItemCount As Long
Dim AparMasSearch As Collection
Dim BdSub As CBillingSubTract
Dim BdAdd As CBillingAddition
Dim Key1 As Long
Dim IsOK As Boolean
Dim Sum8  As Double
Dim Sum9 As Double
Dim Sum10 As Double
Dim Cd  As CConfigDoc

   glbDatabaseMngr.DBConnection.BeginTrans
   HasBegin = True
   Call EnableForm(Me, False)
         
   Set AparMasSearch = New Collection
   I = 0
   prgProgress.Min = 0
   prgProgress.Max = m_BillingDoc.RcpCnDnItems.Count * 2
   
   If Len(m_BillingDoc.NOTE) > 0 Then
      glbDatabaseMngr.DBConnection.RollbackTrans
      glbErrorLog.LocalErrorMsg = "การสร้างเอกสารล้มเหลว เนื่องจากมีการสร้างเอกสารแล้ว"
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   For Each Rcp In m_BillingDoc.RcpCnDnItems
      I = I + 1
      prgProgress.Value = I
      txtPercent.Text = MyDiff(I, m_BillingDoc.RcpCnDnItems.Count) * 50
      Me.Refresh
      
      If Key1 <> Rcp.GetFieldValue("APAR_MAS_ID") Then
         Set TempBd = New CBillingDoc
         Set m_Rs = New ADODB.Recordset
         TempBd.BILLING_DOC_ID = Rcp.GetFieldValue("DOC_ID")
         Call TempBd.QueryData(1, m_Rs, ItemCount)
         Call TempBd.PopulateFromRS(1, m_Rs)
         Key1 = Rcp.GetFieldValue("APAR_MAS_ID")
         
      End If
      
      Set TempBDEX = GetObject("CBillingDoc", AparMasSearch, TempBd.APAR_MAS_ID, False)
      If TempBDEX Is Nothing Then
         TempBd.ShowMode = SHOW_ADD
         TempBd.DOCUMENT_NO = ""            ' ไป SET ตอน Save
         TempBd.DOCUMENT_DATE = DocumentDate            ' เอามาจากหน้าจอ ใบเพิ่มเป็นชุดเลย
         TempBd.DOCUMENT_TYPE = RECEIPT2_DOCTYPE           'เป็นประเภทใบเสร็จรับชำระ
         TempBd.Due_Date = DocumentDate
         TempBd.NOTE = ""           ' ต้อง Clear เป็น ""
         ' BILLING_ADDRESS_ID 'เท่ากับอันเดิม
         ' ENTERPRISE_ADDRESS_ID 'เท่ากับอันเดิม
         TempBd.INVENTORY_DOC_ID = -1             ' ต้อง Clear เป็น -1"
         ' APAR_MAS_ID 'เท่ากับอันเดิม
         TempBd.TOTAL_AMOUNT = 0             ' ต้อง Clear เป็น 0"
         TempBd.TOTAL_PRICE = 0             ' ต้อง Clear เป็น 0"
         TempBd.VAT_PERCENT = 0             ' ต้อง Clear เป็น 0"
         TempBd.VAT_AMOUNT = 0             ' ต้อง Clear เป็น 0"
         TempBd.WH_PERCENT = 0             ' ต้อง Clear เป็น 0"
         TempBd.WH_AMOUNT = 0             ' ต้อง Clear เป็น 0"
         TempBd.DISCOUNT_AMOUNT = 0             ' ต้อง Clear เป็น 0"
         TempBd.EXT_DISCOUNT_AMOUNT = 0             ' ต้อง Clear เป็น 0"
         TempBd.EXT_DISCOUNT_AMOUNT = 0             ' ต้อง Clear เป็น 0"
         
         'SALE BY
         'LOCATION_SALE
         'CUS_PO
         'DO_ADDRESS
         'BRANCH_ADDRESS
         TempBd.DOCUMENT_SUB_TYPE = -1
         TempBd.CANCEL_FLAG = "N"
         TempBd.SR_REF_DO_ID = -1
         TempBd.SR_REF_DO_NO = ""
         
         TempBd.BILLING_DOC_PACK = m_BillingDoc.BILLING_DOC_ID
         
         TempBd.PAY_AMOUNT = 0   'เซ็ตให้เป็น 0 ก่อนแล้วค่อย อัพจากค่าด้านล่าง
         TempBd.PAID_AMOUNT = 0  'เซ็ตให้เป็น 0 ก่อนแล้วค่อย อัพจากค่าด้านล่าง
         TempBd.DEBIT_AMOUNT = 0 'เซ็ตให้เป็น 0 ก่อนแล้วค่อย อัพจากค่าด้านล่าง
         TempBd.CREDIT_AMOUNT = 0   'เซ็ตให้เป็น 0 ก่อนแล้วค่อย อัพจากค่าด้านล่าง
         TempBd.CREDIT_AMOUNT = 0   'เซ็ตให้เป็น 0 ก่อนแล้วค่อย อัพจากค่าด้านล่าง
         
         If Rcp.GetFieldValue("DOC_ID_TYPE") = INVOICE_DOCTYPE Then
            TempBd.PAY_AMOUNT = Rcp.GetFieldValue("ITEM_AMOUNT")
            TempBd.PAID_AMOUNT = Rcp.GetFieldValue("PAID_AMOUNT")
         ElseIf Rcp.GetFieldValue("DOC_ID_TYPE") = DN_DOCTYPE Then
             TempBd.DEBIT_AMOUNT = Rcp.GetFieldValue("ITEM_AMOUNT")
         ElseIf Rcp.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Then
             TempBd.CREDIT_AMOUNT = Rcp.GetFieldValue("ITEM_AMOUNT")
         ElseIf Rcp.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Then
             TempBd.CREDIT_AMOUNT = Rcp.GetFieldValue("ITEM_AMOUNT")
         End If
         
         Call AparMasSearch.add(TempBd, Trim(Str(TempBd.APAR_MAS_ID)))
         
         Rcp.Flag = "A"
         Call TempBd.RcpCnDnItems.add(Rcp)
      Else
         If Rcp.GetFieldValue("DOC_ID_TYPE") = INVOICE_DOCTYPE Then
            TempBDEX.PAY_AMOUNT = TempBDEX.PAY_AMOUNT + Rcp.GetFieldValue("ITEM_AMOUNT")
            TempBDEX.PAID_AMOUNT = TempBDEX.PAID_AMOUNT + Rcp.GetFieldValue("PAID_AMOUNT")
         ElseIf Rcp.GetFieldValue("DOC_ID_TYPE") = DN_DOCTYPE Then
             TempBDEX.DEBIT_AMOUNT = TempBDEX.DEBIT_AMOUNT + Rcp.GetFieldValue("ITEM_AMOUNT")
         ElseIf Rcp.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Then
             TempBDEX.CREDIT_AMOUNT = TempBDEX.CREDIT_AMOUNT + Rcp.GetFieldValue("ITEM_AMOUNT")
         ElseIf Rcp.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Then
             TempBDEX.CREDIT_AMOUNT = TempBDEX.CREDIT_AMOUNT + Rcp.GetFieldValue("ITEM_AMOUNT")
         End If
         Rcp.Flag = "A"
         Call TempBd.RcpCnDnItems.add(Rcp)
      End If
   Next Rcp
   
   Set TempBd = Nothing
   
   Dim TempSub As CBillingSubTract
   Dim TempDif As Double
   Dim TempDif2 As Double
   Dim TempDif3 As Double
   Dim TempDif4 As Double
   Dim TempDif5 As Double
   Dim TempDif6 As Double
   Dim CountChk As Long
   For Each BdSub In m_BillingDoc.BillingSubTracts
      TempDif = 0
      CountChk = 0
      For Each TempBd In AparMasSearch
         CountChk = CountChk + 1
         Set TempSub = New CBillingSubTract
         TempSub.Flag = "A"
         Call TempSub.SetFieldValue("SUBTRACT_ID", BdSub.GetFieldValue("SUBTRACT_ID"))
         Call TempSub.SetFieldValue("ITEM_AMOUNT", Format((TempBd.PAID_AMOUNT + TempBd.DEBIT_AMOUNT + TempBd.CREDIT_AMOUNT) / (m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT + m_BillingDoc.CREDIT_AMOUNT) * BdSub.GetFieldValue("ITEM_AMOUNT"), "0.00"))
         
         TempDif = TempDif + Round((TempBd.PAID_AMOUNT + TempBd.DEBIT_AMOUNT + TempBd.CREDIT_AMOUNT) / (m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT + m_BillingDoc.CREDIT_AMOUNT) * BdSub.GetFieldValue("ITEM_AMOUNT"), 2)
         
         If CountChk = AparMasSearch.Count Then
            If Not (Round(TempDif, 2) = Round(BdSub.GetFieldValue("ITEM_AMOUNT"), 2)) Then
               Call TempSub.SetFieldValue("ITEM_AMOUNT", TempSub.GetFieldValue("ITEM_AMOUNT") - (Round(TempDif, 2) - Round(BdSub.GetFieldValue("ITEM_AMOUNT"), 2)))
            End If
         End If
         
         Call TempBd.BillingSubTracts.add(TempSub)
         Set TempSub = Nothing
      Next TempBd
      
   Next BdSub
   'debug.print
   Set TempBd = Nothing
   
   Dim TempAdd As CBillingAddition
   For Each BdAdd In m_BillingDoc.BillingAdditions
      TempDif = 0
      CountChk = 0
      For Each TempBd In AparMasSearch
         CountChk = CountChk + 1
         Set TempAdd = New CBillingAddition
         TempAdd.Flag = "A"
         Call TempAdd.SetFieldValue("ADDITION_ID", BdAdd.GetFieldValue("ADDITION_ID"))
         
         Call TempAdd.SetFieldValue("ITEM_AMOUNT", Format((TempBd.PAID_AMOUNT + TempBd.DEBIT_AMOUNT + TempBd.CREDIT_AMOUNT) / (m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT + m_BillingDoc.CREDIT_AMOUNT) * BdAdd.GetFieldValue("ITEM_AMOUNT"), "0.00"))
         
         TempDif = TempDif + Round((TempBd.PAID_AMOUNT + TempBd.DEBIT_AMOUNT + TempBd.CREDIT_AMOUNT) / (m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT + m_BillingDoc.CREDIT_AMOUNT) * BdAdd.GetFieldValue("ITEM_AMOUNT"), 2)
         
         If CountChk = AparMasSearch.Count Then
            If Not (Round(TempDif, 2) = Round(BdAdd.GetFieldValue("ITEM_AMOUNT"), 2)) Then
               Call TempAdd.SetFieldValue("ITEM_AMOUNT", TempAdd.GetFieldValue("ITEM_AMOUNT") - (Round(TempDif, 2) - Round(BdAdd.GetFieldValue("ITEM_AMOUNT"), 2)))
            End If
         End If
         
         Call TempBd.BillingAdditions.add(TempAdd)
         Set TempSub = Nothing
      Next TempBd
   Next BdAdd
   Set TempBd = Nothing
   
   Dim Sum1 As Double
   
   Sum1 = 0
   For Each TempBd In AparMasSearch
      I = I + 1
      prgProgress.Value = I
      txtPercent.Text = 70 + (I / AparMasSearch.Count) * 30
      Me.Refresh
         
      Sum1 = 0
      For Each BdSub In TempBd.BillingSubTracts
         Sum1 = Sum1 + BdSub.GetFieldValue("ITEM_AMOUNT")
      Next BdSub
      TempBd.SUBTRACT_AMOUNT = Sum1
      
      Sum1 = 0
      For Each BdAdd In TempBd.BillingAdditions
         Sum1 = Sum1 + BdAdd.GetFieldValue("ITEM_AMOUNT")
      Next BdAdd
      TempBd.ADDITION_AMOUNT = Sum1
   Next TempBd
   prgProgress.Value = I
   txtPercent.Text = 60
   Me.Refresh
   
   Dim BDTran As CCashTran
   Dim TempTran As CCashTran
   For Each BDTran In m_BillingDoc.Payments
      TempDif = 0
      TempDif2 = 0
      TempDif3 = 0
      TempDif4 = 0
      TempDif5 = 0
      TempDif6 = 0
      CountChk = 0
      For Each TempBd In AparMasSearch
         CountChk = CountChk + 1
                  
         Set TempTran = New CCashTran
         TempTran.Flag = "A"
         
         Call TempTran.SetFieldValue("CHEQUE_ID", BDTran.GetFieldValue("CHEQUE_ID"))
         Call TempTran.SetFieldValue("BANK_ID", BDTran.GetFieldValue("BANK_ID"))
         Call TempTran.SetFieldValue("BANK_BRANCH", BDTran.GetFieldValue("BANK_BRANCH"))
         Call TempTran.SetFieldValue("TX_TYPE", BDTran.GetFieldValue("TX_TYPE"))
         Call TempTran.SetFieldValue("PAYMENT_TYPE", BDTran.GetFieldValue("PAYMENT_TYPE"))
         Call TempTran.SetFieldValue("BANK_ACCOUNT", BDTran.GetFieldValue("BANK_ACCOUNT"))
         Call TempTran.SetFieldValue("TX_NO", BDTran.GetFieldValue("TX_NO"))
         Call TempTran.SetFieldValue("TX_DATE", BDTran.GetFieldValue("TX_DATE"))
         Call TempTran.SetFieldValue("APAR_MAS_ID", BDTran.GetFieldValue("APAR_MAS_ID"))
         Call TempTran.SetFieldValue("EMP_ID", BDTran.GetFieldValue("EMP_ID"))
         Call TempTran.SetFieldValue("STEP_ID", BDTran.GetFieldValue("STEP_ID"))
         
         Call TempTran.SetFieldValue("FEE_AMOUNT", Format((TempBd.PAID_AMOUNT + TempBd.DEBIT_AMOUNT + TempBd.CREDIT_AMOUNT) / (m_BillingDoc.PAID_AMOUNT + m_BillingDoc.DEBIT_AMOUNT + m_BillingDoc.CREDIT_AMOUNT) * BDTran.GetFieldValue("FEE_AMOUNT"), "0.00"))
         
         Call TempTran.SetFieldValue("AMOUNT", TempBd.PAID_AMOUNT + TempBd.DEBIT_AMOUNT - TempBd.CREDIT_AMOUNT - TempBd.SUBTRACT_AMOUNT + TempBd.ADDITION_AMOUNT)
         
         Call TempTran.SetFieldValue("NET_AMOUNT", TempTran.GetFieldValue("AMOUNT") - TempTran.GetFieldValue("FEE_AMOUNT"))
         
         Call TempTran.Cheque.SetFieldValue("BANK_ID", BDTran.Cheque.GetFieldValue("BANK_ID"))
         Call TempTran.Cheque.SetFieldValue("BANK_BRANCH", BDTran.Cheque.GetFieldValue("BANK_BRANCH"))
         Call TempTran.Cheque.SetFieldValue("DIRECTION", BDTran.Cheque.GetFieldValue("DIRECTION"))
         Call TempTran.Cheque.SetFieldValue("CHEQUE_NO", BDTran.Cheque.GetFieldValue("CHEQUE_NO"))
         Call TempTran.Cheque.SetFieldValue("CHEQUE_AMOUNT", TempTran.GetFieldValue("AMOUNT"))
         Call TempTran.Cheque.SetFieldValue("CHEQUE_DATE", BDTran.Cheque.GetFieldValue("CHEQUE_DATE"))
         Call TempTran.Cheque.SetFieldValue("EFFECTIVE_DATE", BDTran.Cheque.GetFieldValue("EFFECTIVE_DATE"))
         Call TempTran.Cheque.SetFieldValue("CHEQUE_TYPE", BDTran.Cheque.GetFieldValue("CHEQUE_TYPE"))
         Call TempTran.Cheque.SetFieldValue("APAR_MAS_ID", BDTran.Cheque.GetFieldValue("APAR_MAS_ID"))
         Call TempTran.Cheque.SetFieldValue("CHEQUE_STATUS", BDTran.Cheque.GetFieldValue("CHEQUE_STATUS"))
         Call TempTran.Cheque.SetFieldValue("BANK_FLAG", BDTran.Cheque.GetFieldValue("BANK_FLAG"))
         Call TempTran.Cheque.SetFieldValue("POST_FLAG", BDTran.Cheque.GetFieldValue("POST_FLAG"))
         
         TempDif = TempDif + TempTran.GetFieldValue("AMOUNT")
         TempDif2 = TempDif2 + TempTran.GetFieldValue("FEE_AMOUNT")
         TempDif3 = TempDif3 + TempTran.GetFieldValue("NET_AMOUNT")
         
         If CountChk = AparMasSearch.Count Then
            If Not (Round(TempDif, 2) = Round(BDTran.GetFieldValue("AMOUNT"), 2)) Then
               Call TempTran.SetFieldValue("AMOUNT", TempTran.GetFieldValue("AMOUNT") - (Round(TempDif, 2) - Round(BDTran.GetFieldValue("AMOUNT"), 2)))
            End If
            If Not (Round(TempDif2, 2) = Round(BDTran.GetFieldValue("FEE_AMOUNT"), 2)) Then
               Call TempTran.SetFieldValue("FEE_AMOUNT", TempTran.GetFieldValue("FEE_AMOUNT") - (Round(TempDif2, 2) - Round(BDTran.GetFieldValue("FEE_AMOUNT"), 2)))
            End If
            If Not (Round(TempDif3, 2) = Round(BDTran.GetFieldValue("NET_AMOUNT"), 2)) Then
               Call TempTran.SetFieldValue("NET_AMOUNT", TempTran.GetFieldValue("NET_AMOUNT") - (Round(TempDif3, 2) - Round(BDTran.GetFieldValue("NET_AMOUNT"), 2)))
            End If
         End If
         
         Call TempBd.Payments.add(TempTran)
         Set TempTran = Nothing
      Next TempBd
   Next BDTran
   
   Set TempBd = Nothing
   
   prgProgress.Value = I
   txtPercent.Text = 70
   Me.Refresh
   
   I = 0
   prgProgress.Max = AparMasSearch.Count
   Dim ConFigDocType As Long
   Dim RunningNo As Long
   Dim HeadNo As String
   Dim TempStr As String
   Dim Sum2 As Double
   Dim Sum3 As Double
   Dim Sum4 As Double
   Dim Sum5 As Double
   Dim Sum6 As Double
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   Sum4 = 0
   Sum5 = 0
   Sum6 = 0
   For Each TempBd In AparMasSearch
      I = I + 1
      prgProgress.Value = I
      txtPercent.Text = 70 + (I / AparMasSearch.Count) * 30
      Me.Refresh
      If I = 1 Then
         TempBd.DOCUMENT_NO = GetDocumentNo(RECEIPT2_DOCTYPE, -1, m_BillingDoc.DOCUMENT_DATE, HeadNo, RunningNo, ConFigDocType, TempStr)
         m_BillingDoc.NOTE = TempBd.DOCUMENT_NO
      ElseIf I = AparMasSearch.Count Then
         RunningNo = RunningNo + 1
         TempBd.DOCUMENT_NO = HeadNo & Format(Trim(Str(RunningNo)), TempStr)
         m_BillingDoc.NOTE = m_BillingDoc.NOTE & "-" & RunningNo
         Call m_BillingDoc.UpdateNote
         
         Set Cd = New CConfigDoc
         Call Cd.SetFieldValue("RUNNING_NO", RunningNo)
         Call Cd.SetFieldValue("LAST_NO", HeadNo & Format(Trim(Str(RunningNo)), TempStr))
         Call Cd.SetFieldValue("CONFIG_DOC_TYPE", ConFigDocType)
         Call Cd.UpdateRunningNo
      Else
         RunningNo = RunningNo + 1
         TempBd.DOCUMENT_NO = HeadNo & Format(Trim(Str(RunningNo)), TempStr)
      End If
      
      Sum1 = 0
      Sum2 = 0
      Sum3 = 0
      Sum4 = 0
      Sum5 = 0
      Sum6 = 0
      For Each BDTran In TempBd.Payments
         If BDTran.GetFieldValue("PAYMENT_TYPE") = 1 Then
            Sum1 = Sum1 + BDTran.GetFieldValue("AMOUNT")
         End If
         If BDTran.GetFieldValue("PAYMENT_TYPE") = 2 Or BDTran.GetFieldValue("PAYMENT_TYPE") = 3 Then
            Sum2 = Sum2 + BDTran.GetFieldValue("AMOUNT")
         End If
         If BDTran.GetFieldValue("PAYMENT_TYPE") = 4 Then
            Sum3 = Sum3 + BDTran.GetFieldValue("AMOUNT")
         End If
         
         Sum4 = Sum4 + BDTran.GetFieldValue("NET_AMOUNT")
         Sum5 = Sum5 + BDTran.GetFieldValue("FEE_AMOUNT")
         
      Next BDTran
      TempBd.CASH_PMT = Sum1
      TempBd.CHEQUE_PMT = Sum2
      TempBd.BANKTRF_PMT = Sum3
      TempBd.RCP_CASH_TRAN = Sum4         'หักค่าธรรมเนียมแล้ว
      TempBd.FEE_AMOUNT = Sum5
      
      Call glbDaily.AddEditBillingDoc(TempBd, IsOK, False, glbErrorLog)
   Next TempBd
   
   Set TempBd = Nothing
   Set AparMasSearch = Nothing
   
   
   prgProgress.Value = prgProgress.Max
   txtPercent.Text = 100
   Call EnableForm(Me, True)
   glbDatabaseMngr.DBConnection.CommitTrans
   glbErrorLog.LocalErrorMsg = "การสร้างเอกสารสำเร็จ"
   glbErrorLog.ShowUserError
Exit Sub
Errorhanderror:
   If HasBegin Then
      glbDatabaseMngr.DBConnection.RollbackTrans
      glbErrorLog.LocalErrorMsg = "การสร้างเอกสารล้มเหลว"
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
   End If
End Sub
Private Sub Form_Activate()
Dim FromDate As Date
Dim ToDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadConfigDoc(Nothing, m_Cd)
      
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
   pnlHeader.Caption = "สร้างเอกสารใบเสร็จรับชำระเป็นชุด"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   
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
   m_HasActivate = False
   
   Set m_Cd = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
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

Private Sub Form_Unload(Cancel As Integer)
   Set m_Cd = Nothing
End Sub
