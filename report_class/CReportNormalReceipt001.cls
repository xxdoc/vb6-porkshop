VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportNormalRcp001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportNormalRcp001"
Private MODULE_DESC  As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset

Private Const ENTERPRISE_SIZE = 16
Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 14
Private Const DETAIL_SIZE = 14

Private Const TITLE_FORMAT = "<52|<3|<5;"

Private Const HEADER_RECEIPT1 = "^1|^7|^2|^2|^2;"
Private Const HEADER_RECEIPT2 = "^1|^3|^3|^3|^3|^3|^3;"
Private Const HEADER_RECEIPT3 = "^1|^2|^8|^2|^2|^2|^2;"
Private Const HEADER_RECEIPT4 = "^1|^5|^2|^2|^2|^2;"
Private Const HEADER_BILLS = "^1|^3|^3|^3|^3|^3|^3;"
Private Const HEADER_CNDN = "^1|^4|^4|^4;"

Private Const DETAIL_RECEIPT1 = "^0|^1|<7|>2|>2|>2|^0|^0;"
Private Const DETAIL_RECEIPT2 = "^1|^3|^3|^3|>3|>3|>3;"
Private Const DETAIL_RECEIPT3 = "^1|^2|<8|^2|>2|>2|>2;"
Private Const DETAIL_RECEIPT4 = "^0|^1|<5|>2|>2|>2|>2|^0|^0;"
Private Const DETAIL_BILLS = "^1|^3|^3|^3|>3|>3|>3;"
Private Const DETAIL_CNDN = "^1|^4|^4|>4;"

Private Const FOOTER_RECEIPT1 = "<10|^2|>2;"
Private Const FOOTER_RECEIPT1_1 = "<10|>2|>2;"
Private Const FOOTER_RECEIPT1_2 = "<8|>4|>2;"
Private Const FOOTER_RECEIPT1_3 = ">12|>2;"
Private Const FOOTER_RECEIPT2 = "<13|^3|>3;"
Private Const FOOTER_BILLS = "<13|^3|>3;"
Private Const FOOTER_CNDN = "<5|>4|>4;"

Private m_Enterprise As CEnterprise
Private m_CustAddress As CAddress
Private m_EnpAddress As CAddress
Private m_EnpPerson As CName
Private m_Addr As CAddress
Private m_cData As CBillingDoc
Private m_ChildData As CBillingDoc
Private m_OwnerAddress As CAddress

Private Sp As CSystemParam
Private X1 As Long
Private X2  As Long
Private Y1  As Long
Private Y2  As Long
Private AllEditEndPage As Boolean
Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_Enterprise = New CEnterprise
   Set m_CustAddress = New CAddress
   Set m_EnpAddress = New CAddress
   Set m_EnpPerson = New CName
   Set m_Addr = New CAddress
   Set m_OwnerAddress = New CAddress
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_Enterprise = Nothing
   Set m_CustAddress = Nothing
   Set m_EnpAddress = Nothing
   Set m_EnpPerson = Nothing
   Set m_Addr = Nothing
   Set m_OwnerAddress = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property
Private Function CReportInterface_Preview() As Boolean
   If mcolParam("PREVIEW_TYPE") = 2 Then
      CReportInterface_Preview = genDoc(False)
   Else
      CReportInterface_Preview = genDoc(True)
   End If
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

End Sub
'============================= Codes above used in every report ==========================

Private Sub PrintBillFooter(alngX() As Long)
Dim strFormat As String
Dim strPrint As String

   Vsp.TableBorder = tbNone
   Vsp.FontBold = False
   
   If mcolParam("DOCUMENT_TYPE") <> S_PO_DOCTYPE And mcolParam("DOCUMENT_TYPE") <> INVOICE_DOCTYPE Then
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   
   If mcolParam("DOCUMENT_TYPE") = BILLS_DOCTYPE Then
   
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "หมายเหตุ  " & m_cData.NOTE & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Call Vsp.AddTable(strFormat, "", strPrint)
      Call Vsp.AddTable(strFormat, "", strPrint)
      
   
      strFormat = VSP_CalTable("<3|<7|<10;", mdblWid, alngX)
      strPrint = "ชื่อผู้รับวางบิล|........./........./.........|ในนาม " & m_Enterprise.ENTERPRISE_NAME & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<3|<7|10;", mdblWid, alngX)
      strPrint = "วันที่รับ|........./........./.........|;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<3|<7|<10;", mdblWid, alngX)
      strPrint = " วันที่นัดรับเช็ค|........./........./.........|ชื่อผู้วางบิล .....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = RECEIPT1_DOCTYPE Then
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "ชื่อผู้รับเงิน|.....................................|ชื่อผู้จ่ายเงิน|.....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "วันที่|........./........./.........|วันที่|........./........./.........;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = RECEIPT2_DOCTYPE Then
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "การชำระเงินด้วยเช็คจะสมบูรณ์เมื่อบริษัทได้รับเงินตามเช็คเรียบร้อย;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "เงินสด .......................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<10|<10|<10;", mdblWid, alngX)
      strPrint = "เช็คธนาคาร ............................|เช็คเลขที่ ..................................|ลงวันที่ ........./........./.........|จำนวนเงิน ..................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<10|<10|<10;", mdblWid, alngX)
      strPrint = "เช็คธนาคาร ............................|เช็คเลขที่ ..................................|ลงวันที่ ........./........./.........|จำนวนเงิน ..................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "ชื่อผู้รับเงิน|.....................................|ชื่อผู้จ่ายเงิน|.....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "วันที่|........./........./.........|วันที่|........./........./.........;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = S_RECEIPT2_DOCTYPE Then
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "ผู้จัดทำ|.....................................|ผู้อนุมัติ|.....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<3|<7;", mdblWid, alngX)
      strPrint = "วันที่|........./........./.........|วันที่|........./........./.........;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
   ElseIf mcolParam("DOCUMENT_TYPE") = S_PO_DOCTYPE Then
      strFormat = VSP_CalTable("^10;", mdblWid, alngX)
      strPrint = "ในนาม" & m_Enterprise.ENTERPRISE_NAME & " ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10|^10|^10;", mdblWid, alngX)
      strPrint = "............................|..................................|..................................|..................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10|^10|^10;", mdblWid, alngX)
      strPrint = "ผู้จัดทำ|ผู้สั่งซื้อ|ผู้ตรวจสอบ|ผู้อนุมัติ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      AllEditEndPage = False
   ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      Vsp.FontBold = True
      strPrint = "โปรดสั่งจ่ายเช็คขีดคร่อมในนาม " & m_Enterprise.ENTERPRISE_NAME & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "บริษัทฯ ขอสงวนสิทธิ์ที่จะคิดดอกเบี้ยในอัตรา 1.5% ต่อเดือน ในกรณีที่ท่านชำระเงินเกินกำหนดเวลา;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Vsp.FontBold = False
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "ได้รับสินค้าตามรายการข้างบนนี้ไว้ถูกต้องและอยู่ในสภาพเรียบร้อย;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "Received as per above goods in good order and condition;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10|^10|^10|^10;", mdblWid, alngX)
      strPrint = "............................|..................................|..................................|..................................|..................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10|^10|^10|^10;", mdblWid, alngX)
      strPrint = "ผู้รับสินค้า|วันที่รับสินค้า|ผู้ส่งสินค้า|ผู้ตรวจสอบ|ผู้รับมอบอำนาจ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10|^10|^10|^10;", mdblWid, alngX)
      strPrint = "Received by|Date|Delivered by|Approved by|Authorized Sign;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      AllEditEndPage = False
   ElseIf mcolParam("DOCUMENT_TYPE") = CN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = DN_DOCTYPE Then
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "หมายเหตุ  " & m_cData.NOTE & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strPrint = ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Call Vsp.AddTable(strFormat, "", strPrint)
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<10|<10;", mdblWid, alngX)
      strPrint = " ผู้จัดทำ.....................................|ผู้รับมอบอำนาจ .....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_INVOICE_DOCTYPE Then
      strFormat = VSP_CalTable("^10|^10|^5|^10;", mdblWid, alngX)
      strPrint = "ได้รับสินค้าข้างต้นเรียบร้อยแล้ว|ได้เก็บสินค้าข้างต้นในคลังสินค้าแล้ว||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp = ""
      strPrint = ".....................................|.....................................|.............................|.....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "ผู้จัดทำ|แผนกคลังสินค้า|ผู้ตรวจสอบ|ผู้รับมอบอำนาจ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
  ElseIf mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Then
      strFormat = VSP_CalTable("^10|^10|^5|^10;", mdblWid, alngX)
      strPrint = "ได้รับสินค้าข้างต้นเรียบร้อยแล้ว|ได้ส่งคืนสินค้าข้างต้นเรียบร้อยแล้ว||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp = ""
      strPrint = ".....................................|.....................................|.............................|.....................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = "ผู้รับสินค้าคืน|แผนกคลังสินค้า|ผู้ตรวจสอบ|ผู้รับมอบอำนาจ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
      strFormat = VSP_CalTable("^10;", mdblWid, alngX)
      strPrint = "ทางเราขอขอบพระคุณที่ให้ความไว้วางใจด้วยดีเสมอมา;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strPrint = ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      strPrint = "..................................|..................................;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("^10|^10;", mdblWid, alngX)
      strPrint = "ผู้จัดทำ|ผู้อนุมัติ;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      AllEditEndPage = False
   End If
End Sub

Private Sub PrintBillHeader(alngX() As Long)
Dim strFormat As String
Dim strPrint As String
Static I As Long
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim OwnerFlag As String
Dim TempStr As String
Dim TempAddr As String
Dim Doc As CDocItem
   
   If I = 0 Then
      I = 1
      Set TempRs = New ADODB.Recordset
      
      m_Enterprise.ENTERPRISE_ID = -1
      Call m_Enterprise.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_Enterprise.PopulateFromRS(1, TempRs)
      End If
      
      Call m_EnpAddress.SetFieldValue("ADDRESS_ID", m_cData.ENTERPRISE_ADDRESS_ID)
      Call m_EnpAddress.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call m_EnpAddress.PopulateFromRS(1, TempRs)
      End If
      If m_cData.BILLING_ADDRESS_ID > 0 Then
         Call m_CustAddress.SetFieldValue("ADDRESS_ID", m_cData.BILLING_ADDRESS_ID)
         Call m_CustAddress.QueryData(1, TempRs, iCount)
         If Not TempRs.EOF Then
            Call m_CustAddress.PopulateFromRS(1, TempRs)
         End If
      End If
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   Vsp.TableBorder = tbNone
   Vsp.FontSize = ENTERPRISE_SIZE
   Vsp.FontBold = True
   
   If mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      strPrint = m_Enterprise.ENTERPRISE_NAME & "|" & "ใบเสนอราคา/ใบสั่งซื้อ" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
         
      Vsp.FontSize = TITLE_SIZE
      
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & "|" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = "โทร. " & EmptyToString(m_EnpAddress.GetFieldValue("PHONE1"), "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.GetFieldValue("FAX1"), "N/A") & "|" & "เลขที่" & "|" & m_cData.DOCUMENT_NO & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = "|" & "วันที่|" & DateToStringExtEx2((m_cData.DOCUMENT_DATE)) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<2|<10|;", mdblWid, alngX)
      strPrint = "ลูกค้า| " & m_cData.APAR_CODE & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<2|<10;", mdblWid, alngX)
      strPrint = "|" & m_cData.APAR_NAME & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = S_PO_DOCTYPE Then
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      strPrint = m_Enterprise.ENTERPRISE_NAME & "|" & "ใบสั่งซื้อ" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      Vsp.FontSize = TITLE_SIZE
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      strPrint = "โทร. " & EmptyToString(m_EnpAddress.GetFieldValue("PHONE1"), "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.GetFieldValue("FAX1"), "N/A") & "|" & SellDoctype2Report(mcolParam("DOCUMENT_TYPE")) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & m_Enterprise.TAX_ID & "|" & "เลขที่" & SellDoctype2Report(mcolParam("DOCUMENT_TYPE")) & "|" & m_cData.DOCUMENT_NO & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
      strPrint = "ผู้จำหน่าย| " & m_cData.APAR_CODE & "|" & "วันที่|" & DateToStringExtEx2((m_cData.DOCUMENT_DATE)) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = m_cData.APAR_NAME & "|" & "วันที่ส่งของ|" & DateToStringExtEx2(m_cData.Due_Date) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      strPrint = m_Enterprise.ENTERPRISE_NAME & "|" & "" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
              
      Vsp.FontSize = TITLE_SIZE
         
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & "|" & "ใบส่งของ/ใบกำกับภาษี" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = "โทร. " & EmptyToString(m_EnpAddress.GetFieldValue("PHONE1"), "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.GetFieldValue("FAX1"), "N/A") & "|" & "เลขที่" & SellDoctype2Report(mcolParam("DOCUMENT_TYPE")) & "|" & m_cData.DOCUMENT_NO & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & m_Enterprise.TAX_ID & "|" & "วันที่|" & DateToStringExtEx2((m_cData.DOCUMENT_DATE)) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
      strPrint = "ชื่อลูกค้า| " & m_cData.APAR_CODE & "|" & "วันที่ครบกำหนด|" & DateToStringExtEx2(m_cData.Due_Date) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = m_cData.APAR_NAME & "|" & "เครดิต" & "|" & m_cData.CREDIT & " วัน" & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   Else
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = m_Enterprise.ENTERPRISE_NAME & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)

      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = MapText("ที่อยู่ ") & m_EnpAddress.PackAddress & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10|^10;", mdblWid, alngX)
      If 6 = mcolParam("DOCUMENT_TYPE") Then
         Dim m_MasterRef As CMasterRef
         Set m_MasterRef = New CMasterRef
         Set TempRs = New ADODB.Recordset
         
         m_MasterRef.KEY_ID = -1
         m_MasterRef.KEY_CODE = "A-01"
         m_MasterRef.MASTER_AREA = 999
         Call m_MasterRef.QueryData(1, TempRs, iCount, True)
         If Not TempRs.EOF Then
            Call m_MasterRef.PopulateFromRS(1, TempRs)
            TempStr = m_MasterRef.KEY_NAME
         End If
         Set TempRs = Nothing
      End If
      strPrint = "โทร. " & EmptyToString(m_EnpAddress.GetFieldValue("PHONE1"), "N/A") & " แฟกซ์ " & EmptyToString(m_EnpAddress.GetFieldValue("FAX1"), "N/A") & "|" & SellDoctype2Report(mcolParam("DOCUMENT_TYPE")) & TempStr & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = MapText("เลขประจำตัวผู้เสียภาษี ") & m_Enterprise.TAX_ID & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)

      strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
      If mcolParam("DOCUMENT_TYPE") <= 100 Then
         strPrint = "ลูกค้า| " & m_cData.APAR_CODE & "|" & "เลขที่" & SellDoctype2Report(mcolParam("DOCUMENT_TYPE")) & "|" & m_cData.DOCUMENT_NO & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
      Else
         strPrint = "ผู้ขาย| " & m_cData.APAR_CODE & "|" & "เลขที่" & SellDoctype2Report(mcolParam("DOCUMENT_TYPE")) & "|" & m_cData.DOCUMENT_NO & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
      End If
      
      strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
      strPrint = m_cData.APAR_NAME & "|" & "วันที่|" & DateToStringExtEx2((m_cData.DOCUMENT_DATE)) & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
   End If
   strFormat = VSP_CalTable("<10|<6|<4;", mdblWid, alngX)
   If mcolParam("DOCUMENT_TYPE") = BILLS_DOCTYPE Then
      strPrint = EmptyToString(m_CustAddress.PackAddress, "N/A") & "|เงื่อนไขการชำระเงิน|" & DateToStringExtEx2(m_cData.Due_Date) & ";"
   ElseIf mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Then
      Set Doc = m_cData.DocItems.Item(1)
      strPrint = EmptyToString(m_CustAddress.PackAddress, "N/A") & "|อ้างถึงใบกำกับภาษี|" & Doc.GetFieldValue("PO_NO") & ";"
   ElseIf mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Then
      Set Doc = m_cData.DocItems.Item(1)
      strPrint = EmptyToString(m_CustAddress.PackAddress, "N/A") & "|อ้างถึงใบรับสินค้า|" & Doc.GetFieldValue("PO_NO") & ";"
   ElseIf mcolParam("DOCUMENT_TYPE") = S_INVOICE_DOCTYPE Then
      strPrint = EmptyToString(m_CustAddress.PackAddress, "N/A") & "|อ้างถึงใบสั่งซื้อ|" & m_cData.CUS_PO & ";"
   ElseIf mcolParam("DOCUMENT_TYPE") = S_PO_DOCTYPE Then
      strPrint = EmptyToString(m_CustAddress.PackAddress, "N/A") & "|เครดิต|" & m_cData.CREDIT & " วัน" & ";"
   ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
      strPrint = EmptyToString(m_CustAddress.PackAddress, "N/A") & "|รหัสร้านค้า/PO NO.|" & m_cData.CUS_PO & ";"
   ElseIf mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "ที่อยู่ " & EmptyToString(m_CustAddress.PackAddress, "N/A") & ";"
   Else
      strPrint = "ที่อยู่ " & EmptyToString(m_CustAddress.PackAddress, "N/A") & "||;"
   End If
   
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   If mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Then
      Set m_ChildData = New CBillingDoc
      Set TempRs = New ADODB.Recordset
      m_ChildData.BILLING_DOC_ID = Doc.GetFieldValue("PO_ID")
      Call m_ChildData.QueryData(1, TempRs, iCount)

      If iCount > 0 Then
         Call m_ChildData.PopulateFromRS(1, TempRs)
      End If
      
      strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
      strPrint = "TAX ID|" & m_cData.TAX_ID & "|พนักงานขาย|" & m_cData.SALE_CODE & "-" & m_cData.SALE_NAME & ";"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
      strPrint = "|||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "บริษัทได้รับคืนสินค้าและเครดิตบัญชีของท่านตามรายการต่อไปนี้;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   ElseIf mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Then
      Set m_ChildData = New CBillingDoc
      Set TempRs = New ADODB.Recordset
      m_ChildData.BILLING_DOC_ID = Doc.GetFieldValue("PO_ID")
      Call m_ChildData.QueryData(1, TempRs, iCount)

      If iCount > 0 Then
         Call m_ChildData.PopulateFromRS(1, TempRs)
      End If
      
      strFormat = VSP_CalTable("<3|<7|<6|<4;", mdblWid, alngX)
      strPrint = "|||;"
      Call Vsp.AddTable(strFormat, "", strPrint)
'
      strFormat = VSP_CalTable("<10;", mdblWid, alngX)
      strPrint = "บริษัทได้รับคืนสินค้าและเครดิตบัญชีของท่านตามรายการต่อไปนี้;"
      Call Vsp.AddTable(strFormat, "", strPrint)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   Vsp.FontBold = False
End Sub

Private Sub printHeader()
   Dim strFormat As String
   Dim strPrint As String
   Dim tbdt As TableBorderSettings
   Dim blnBold As Boolean
   Dim iSize As Integer
   Dim alngX() As Long
   
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Call PrintBillHeader(alngX)
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
      
   If mcolParam("DOCUMENT_TYPE") = RECEIPT1_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_INVOICE_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_RECEIPT4, mdblWid, alngX)
      strPrint = "No.|รายการ|จำนวน|ราคาต่อหน่วย|ส่วนลด|จำนวนเงิน;"
   ElseIf mcolParam("DOCUMENT_TYPE") = RECEIPT2_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RECEIPT2_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_RECEIPT2, mdblWid, alngX)
      strPrint = "No.|เลขที่ใบกำกับ|วันที่|ครบกำหนด|จำนวนเงิน|ยอดค้าง|ยอดชำระ;"
   ElseIf mcolParam("DOCUMENT_TYPE") = BILLS_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_BILLS_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_BILLS, mdblWid, alngX)
      strPrint = "No.|เลขที่ใบกำกับ|วันที่|ครบกำหนด|จำนวนเงิน|ชำระแล้ว|เงินคงค้าง;"
   ElseIf mcolParam("DOCUMENT_TYPE") = CN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_CN_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_CNDN, mdblWid, alngX)
      strPrint = "No.|เลขที่ใบกำกับ|วันที่|ส่วนลดหนี้;"
   ElseIf mcolParam("DOCUMENT_TYPE") = DN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_DN_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_CNDN, mdblWid, alngX)
      strPrint = "No.|เลขที่ใบกำกับ|วันที่|ส่วนเพิ่มหนี้;"
   ElseIf mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_RECEIPT4, mdblWid, alngX)
      strPrint = "No.|รายการ|จำนวน|ราคาต่อหน่วย|ส่วนลด|จำนวนเงิน;"
   ElseIf mcolParam("DOCUMENT_TYPE") = S_PO_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_RECEIPT2, mdblWid, alngX)
      strPrint = "No.|รหัส|รายละเอียด|จำนวน|ราคา/หน่วย|ส่วนลด|รวม;"
   ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
      strFormat = VSP_CalTable(HEADER_RECEIPT3, mdblWid, alngX)
      strPrint = "No.|รหัส|รายละเอียด|จำนวน|ราคา/หน่วย|ส่วนลด|รวม;"
   End If
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   X1 = Vsp.PageWidth - Vsp.MarginRight
   Y1 = Vsp.CurrentY
   
   
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
   Dim strDate As String
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4
   
   Vsp.Orientation = orPortrait
   Vsp.MarginBottom = 1440
   Vsp.MarginFooter = 1440
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 700
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
   
   Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
   
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim strFormat As String
Dim strPrint As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Bi As CDocItem
Dim Sum1 As Double
Dim Sum2 As Double
Dim Sum3 As Double
Dim DiscountAmount As Double
Dim VatAmount As Double
Dim IncludeDiscountAmount As Double
Dim IncludeVatAmount As Double
Dim TempVatAmount As Double
Dim TempDeposit As Double
Dim TempDiscount As Double
Dim TempCol As Collection
Dim SpaceCount As Long
Dim Ri As CRcpCnDn_Item
Dim Temp1 As Double
Dim Temp2 As Double

   RName = "genDoc"
   
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set m_cData = New CBillingDoc
   m_cData.BILLING_DOC_ID = mcolParam("BILLING_DOC_ID")
   m_cData.QueryFlag = 1
   If Not glbDaily.QueryBillingDoc(m_cData, Rs, iCount, IsOK, glbErrorLog) Then
      mstrErrMsg = "Query Error."
      Exit Function
   End If
   
   If (Rs Is Nothing) Or (iCount < 1) Then
      mstrErrMsg = "ไม่พบข้อมูลที่ต้องการในฐานข้อมูล"
      Exit Function
   End If
   
   Call m_cData.PopulateFromRS(1, Rs)
   
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
   Sum1 = 0
   Sum2 = 0
   Sum3 = 0
   
   Call StartExportFile(Vsp)
   
   Vsp.PenStyle = psDash
   
   Vsp.StartDoc
      
      Vsp.TableBorder = tbAll
      I = 0
      AllEditEndPage = True
      TempDiscount = 0
      TempVatAmount = 0
         
      Vsp.TableBorder = tbColumns
      Dim Addr As CAddress
      If m_cData.BRANCH_ADDRESS > 0 Then
         Set Addr = New CAddress
         Call Addr.SetFieldValue("ADDRESS_ID", m_cData.BRANCH_ADDRESS)
         Call Addr.QueryData(1, Rs, iCount)
         If Not Rs.EOF Then
            Call Addr.PopulateFromRS(1, Rs)
            Call mcolParam.add(Addr.PackAddress, "BRANCH_ADDRESS")
            Call mcolParam.add(Addr.GetFieldValue("PHONE1"), "BRANCH_ADDRESS_PHONE")
             Call mcolParam.add(Addr.GetFieldValue("FAX1"), "BRANCH_ADDRESS_FAX")
            Call mcolParam.add(Addr.GetFieldValue("SHOW_LOCATION_FLAG"), "SHOW_LOCATION_FLAG")
         End If
         Set Addr = Nothing
      Else
         Call mcolParam.add("", "BRANCH_ADDRESS")
         Call mcolParam.add("", "BRANCH_ADDRESS_PHONE")
         Call mcolParam.add("", "BRANCH_ADDRESS_FAX")
         Call mcolParam.add("", "SHOW_LOCATION_FLAG")
      End If
   
      If mcolParam("DOCUMENT_TYPE") = RECEIPT1_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_INVOICE_DOCTYPE Then
         strFormat = VSP_CalTable(DETAIL_RECEIPT4, mdblWid, alngX)
         Set TempCol = m_cData.DocItems
         For Each Bi In TempCol
            I = I + 1
            strPrint = "|"
            strPrint = strPrint & I & "|"
            strPrint = strPrint & Bi.ShowDescCode & " " & Bi.ShowDescText & "|"
            strPrint = strPrint & FormatNumber(MyDiff(Bi.GetFieldValue("ITEM_AMOUNT"), Bi.GetFieldValue("UNIT_MULTIPLE"))) & " " & Bi.GetFieldValue("UNIT_TRAN_NAME") & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("AVG_PRICE") * Bi.GetFieldValue("UNIT_MULTIPLE")) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("DISCOUNT_AMOUNT")) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("TOTAL_PRICE") - Bi.GetFieldValue("DISCOUNT_AMOUNT")) & "|12|12;"
            
            Sum3 = Sum3 + Bi.GetFieldValue("TOTAL_PRICE") - Bi.GetFieldValue("DISCOUNT_AMOUNT")
            TempDiscount = TempDiscount + Bi.GetFieldValue("DISCOUNT_AMOUNT")
            
            If I >= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 5000)
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Bi
         
      ElseIf mcolParam("DOCUMENT_TYPE") = BILLS_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = RECEIPT2_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_BILLS_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RECEIPT2_DOCTYPE Then
         strFormat = VSP_CalTable(DETAIL_RECEIPT2, mdblWid, alngX)
         Set TempCol = m_cData.RcpCnDnItems
         For Each Ri In TempCol
            I = I + 1
            strPrint = I & "|"
            strPrint = strPrint & Ri.GetFieldValue("DOC_NO") & "|"
            strPrint = strPrint & DateToStringExtEx2(Ri.GetFieldValue("DOC_DATE")) & "|"
            strPrint = strPrint & DateToStringExtEx2(Ri.GetFieldValue("DUE_DATE")) & "|"
            If Ri.GetFieldValue("DOC_ID_TYPE") = INVOICE_DOCTYPE Or Ri.GetFieldValue("DOC_ID_TYPE") = S_INVOICE_DOCTYPE Then
               strPrint = strPrint & FormatNumber(Ri.GetFieldValue("TOTAL_PRICE") - Ri.GetFieldValue("DISCOUNT_AMOUNT") - Ri.GetFieldValue("EXT_DISCOUNT_AMOUNT") + Ri.GetFieldValue("VAT_AMOUNT")) & "|"
               strPrint = strPrint & FormatNumber(Ri.GetFieldValue("ITEM_AMOUNT") - Ri.GetFieldValue("PAID_AMOUNT") - Ri.GetFieldValue("PAID_DISCOUNT")) & "|"
               strPrint = strPrint & FormatNumber(Ri.GetFieldValue("PAID_AMOUNT") + Ri.GetFieldValue("PAID_DISCOUNT")) & ";"
               Sum3 = Sum3 + Ri.GetFieldValue("PAID_AMOUNT") + Ri.GetFieldValue("PAID_DISCOUNT")
               TempDiscount = TempDiscount + Ri.GetFieldValue("PAID_DISCOUNT")
            ElseIf Ri.GetFieldValue("DOC_ID_TYPE") = RETURN_DOCTYPE Or Ri.GetFieldValue("DOC_ID_TYPE") = S_RETURN_DOCTYPE Then
               strPrint = strPrint & "-" & FormatNumber(Ri.GetFieldValue("TOTAL_PRICE") - Ri.GetFieldValue("DISCOUNT_AMOUNT") - Ri.GetFieldValue("EXT_DISCOUNT_AMOUNT") + Ri.GetFieldValue("VAT_AMOUNT")) & "|"
               strPrint = strPrint & "0.00|"
               strPrint = strPrint & "-" & FormatNumber(Ri.GetFieldValue("TOTAL_PRICE") - Ri.GetFieldValue("DISCOUNT_AMOUNT") - Ri.GetFieldValue("EXT_DISCOUNT_AMOUNT") + Ri.GetFieldValue("VAT_AMOUNT")) & "|"
               Sum3 = Sum3 - FormatNumber(Ri.GetFieldValue("TOTAL_PRICE") - Ri.GetFieldValue("DISCOUNT_AMOUNT") - Ri.GetFieldValue("EXT_DISCOUNT_AMOUNT") + Ri.GetFieldValue("VAT_AMOUNT"))
            ElseIf Ri.GetFieldValue("DOC_ID_TYPE") = CN_DOCTYPE Or Ri.GetFieldValue("DOC_ID_TYPE") = S_CN_DOCTYPE Then
               strPrint = strPrint & "-" & FormatNumber(Ri.GetFieldValue("PAY_AMOUNT")) & "|"
               strPrint = strPrint & "0.00|"
               strPrint = strPrint & "-" & FormatNumber(Ri.GetFieldValue("PAY_AMOUNT")) & "|"
               Sum3 = Sum3 - FormatNumber(Ri.GetFieldValue("PAY_AMOUNT"))
            ElseIf Ri.GetFieldValue("DOC_ID_TYPE") = DN_DOCTYPE Or Ri.GetFieldValue("DOC_ID_TYPE") = S_DN_DOCTYPE Then
               strPrint = strPrint & FormatNumber(Ri.GetFieldValue("PAY_AMOUNT")) & "|"
               strPrint = strPrint & "0.00|"
               strPrint = strPrint & FormatNumber(Ri.GetFieldValue("PAY_AMOUNT")) & "|"
               Sum3 = Sum3 + FormatNumber(Ri.GetFieldValue("PAY_AMOUNT"))
            End If
            
            
            If I >= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 5000)
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Ri
      ElseIf mcolParam("DOCUMENT_TYPE") = CN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = DN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_CN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_DN_DOCTYPE Then
         strFormat = VSP_CalTable(DETAIL_CNDN, mdblWid, alngX)
         Set TempCol = m_cData.RcpCnDnItems
         For Each Ri In TempCol
            I = I + 1
            strPrint = I & "|"
            strPrint = strPrint & Ri.GetFieldValue("DOC_NO") & "|"
            strPrint = strPrint & DateToStringExtEx2(Ri.GetFieldValue("DOC_DATE")) & "|"
            strPrint = strPrint & FormatNumber(Ri.GetFieldValue("CNDN_AMOUNT")) & ";"
            Sum3 = Sum3 + Ri.GetFieldValue("CNDN_AMOUNT")
   
            If I >= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 4500)
            End If
            
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Ri
      ElseIf mcolParam("DOCUMENT_TYPE") = S_PO_DOCTYPE Then
         strFormat = VSP_CalTable(DETAIL_RECEIPT2, mdblWid, alngX)
         Set TempCol = m_cData.DocItems
         For Each Bi In TempCol
            I = I + 1
            strPrint = I & "|"
            strPrint = strPrint & Bi.ShowDescCode & "|"
            strPrint = strPrint & Bi.ShowDescText & "|"
            strPrint = strPrint & FormatNumber(MyDiff(Bi.GetFieldValue("ITEM_AMOUNT"), Bi.GetFieldValue("UNIT_MULTIPLE")), 0) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("AVG_PRICE") * Bi.GetFieldValue("UNIT_MULTIPLE")) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("DISCOUNT_AMOUNT")) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("TOTAL_PRICE") - Bi.GetFieldValue("DISCOUNT_AMOUNT")) & "|"
            Sum3 = Sum3 + (Bi.GetFieldValue("TOTAL_PRICE"))
            Sum2 = Sum2 + (Bi.GetFieldValue("DISCOUNT_AMOUNT"))
            If I >= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 4500)
            End If
            
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Bi
      ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
         strFormat = VSP_CalTable(DETAIL_RECEIPT3, mdblWid, alngX)
         Set TempCol = m_cData.DocItems
         For Each Bi In TempCol
            I = I + 1
            strPrint = I & "|"
            strPrint = strPrint & Bi.ShowDescCode & "|"
            strPrint = strPrint & Bi.ShowDescText & "|"
            strPrint = strPrint & FormatNumber(MyDiff(Bi.GetFieldValue("ITEM_AMOUNT"), Bi.GetFieldValue("UNIT_MULTIPLE"))) & " " & Bi.GetFieldValue("UNIT_TRAN_NAME") & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("AVG_PRICE") * Bi.GetFieldValue("UNIT_MULTIPLE")) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("DISCOUNT_AMOUNT")) & "|"
            strPrint = strPrint & FormatNumber(Bi.GetFieldValue("TOTAL_PRICE") - Bi.GetFieldValue("DISCOUNT_AMOUNT")) & "|"
            Sum3 = Sum3 + (Bi.GetFieldValue("TOTAL_PRICE"))
            Sum2 = Sum2 + (Bi.GetFieldValue("DISCOUNT_AMOUNT"))
            If I >= TempCol.Count Then
               Call AdjustPage(Vsp, strFormat, strPrint, 4500)
            End If
            
            Call Vsp.AddTable(strFormat, "", strPrint)
         Next Bi
      End If
      
      If mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Then
         Call PatchTable(Vsp, strFormat, "||||;", 6000) '10
      ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Then
         If (Sum2 > 0 Or m_cData.EXT_DISCOUNT_AMOUNT > 0) And m_cData.VAT_AMOUNT > 0 And (mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE) Then
            Call PatchTable(Vsp, strFormat, "||||;", 6000) '10
         Else
            Call PatchTable(Vsp, strFormat, "||||;", 4800) '10
         End If
      ElseIf mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
         Call PatchTable(Vsp, strFormat, "||||;", 5000) '10
      Else
         Call PatchTable(Vsp, strFormat, "||||;", 5000) '10
      End If
      
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
      
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Vsp.TableBorder = tbNone
      Y1 = Vsp.CurrentY
      AllEditEndPage = False
      
      If mcolParam("DOCUMENT_TYPE") = RECEIPT1_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_INVOICE_DOCTYPE Then
         strFormat = VSP_CalTable(FOOTER_RECEIPT1, mdblWid, alngX)
         Vsp.FontBold = True
         strPrint = "หมายเหตุ    " & m_cData.NOTE & "|รวม|" & FormatNumber(Sum3) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
         
         If TempDiscount > 0 Then
            If mcolParam("DOCUMENT_TYPE") <> S_INVOICE_DOCTYPE Then
               strPrint = "|ส่วนลด|" & FormatNumber(TempDiscount) & ";"
               Call Vsp.AddTable(strFormat, "", strPrint)
            End If
         End If
                  
         If m_cData.EXT_DISCOUNT_AMOUNT > 0 Then
            strPrint = "|ส่วนลดพิเศษ|" & FormatNumber(m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
            
         If m_cData.VAT_AMOUNT > 0 And (mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") <> S_RETURN_DOCTYPE) Then
            strPrint = "|ภาษี|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
          If mcolParam("DOCUMENT_TYPE") <> S_INVOICE_DOCTYPE Then
            Temp1 = Sum3 - TempDiscount - m_cData.EXT_DISCOUNT_AMOUNT + m_cData.VAT_AMOUNT
         Else
            Temp1 = Sum3 - m_cData.EXT_DISCOUNT_AMOUNT + m_cData.VAT_AMOUNT
         End If
         
         If mcolParam("DOCUMENT_TYPE") = RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RETURN_DOCTYPE Then
            Temp1 = Sum3 - TempDiscount - m_cData.EXT_DISCOUNT_AMOUNT
            Temp2 = m_ChildData.TOTAL_PRICE - m_ChildData.DISCOUNT_AMOUNT - m_ChildData.EXT_DISCOUNT_AMOUNT
            
            If mcolParam("SHOW_LOCATION_FLAG") = "Y" Then
               strFormat = VSP_CalTable(FOOTER_RECEIPT1_1, mdblWid, alngX)
               strPrint = "สถานที่ส่งของ " & m_cData.APAR_NAME & " " & vbCrLf & mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX") & "|รวมสุทธิ|" & FormatNumber(Temp1) & ";"
               Call Vsp.AddTable(strFormat, "", strPrint)
            Else
               strFormat = VSP_CalTable(FOOTER_RECEIPT1_1, mdblWid, alngX)
               strPrint = "|รวมสุทธิ|" & FormatNumber(Temp1) & ";"
               Call Vsp.AddTable(strFormat, "", strPrint)
            End If
         
            strFormat = VSP_CalTable(FOOTER_RECEIPT1_3, mdblWid, alngX)
            strPrint = "มูลค่าของสินค้าหรือบริการตามใบกำกับภาษีเดิม|" & FormatNumber(Temp2) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
            
            strPrint = "มูลค่าของสินค้าหรือบริการที่ถูกต้อง|" & FormatNumber(Temp2 - Temp1) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
            
            strPrint = "ผลต่าง|" & FormatNumber(Temp1) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
             
            If m_cData.VAT_AMOUNT > 0 Then
               strPrint = "จำนวนภาษีมูลค่าเพิ่ม " & m_cData.VAT_PERCENT & " %|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
               Call Vsp.AddTable(strFormat, "", strPrint)
             End If
            strFormat = VSP_CalTable(FOOTER_RECEIPT1_2, mdblWid, alngX)
            strPrint = "(" & ThaiBaht(Temp1 + m_cData.VAT_AMOUNT) & ")|จำนวนเงินรวมทั้งสิ้น|" & FormatNumber(Temp1 + m_cData.VAT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         Else
            strPrint = "(" & ThaiBaht(Temp1) & ")|รวมสุทธิ|" & FormatNumber(Temp1) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
      ElseIf mcolParam("DOCUMENT_TYPE") = S_PO_DOCTYPE Then
         strFormat = VSP_CalTable("<10|<2|>2;", mdblWid, alngX)
         Vsp.FontBold = True
         If Sum2 > 0 Or m_cData.EXT_DISCOUNT_AMOUNT > 0 Or m_cData.VAT_AMOUNT > 0 Or mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Then
            strPrint = m_cData.NOTE & "|รวมเป็นเงิน|" & FormatNumber(Sum3) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If Sum2 > 0 Then
            strPrint = "|ส่วนลด|" & FormatNumber(Sum2) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
                  
         If m_cData.EXT_DISCOUNT_AMOUNT > 0 Then
            strPrint = "|ส่วนลดพิเศษ|" & FormatNumber(m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If (Sum2 > 0 Or m_cData.EXT_DISCOUNT_AMOUNT > 0) And m_cData.VAT_AMOUNT > 0 And mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Then
            strPrint = "|รวมหักส่วนลด|" & FormatNumber(Sum3 - Sum2 - m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If m_cData.VAT_AMOUNT > 0 And mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Then
            strPrint = "|ภาษี " & m_cData.VAT_PERCENT & " %" & "|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         Temp1 = Sum3 - Sum2 - m_cData.EXT_DISCOUNT_AMOUNT + m_cData.VAT_AMOUNT
      
         strPrint = "(" & ThaiBaht(Temp1) & ")|รวมสุทธิ|" & FormatNumber(Temp1) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
      ElseIf mcolParam("DOCUMENT_TYPE") = INVOICE_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = PO_DOCTYPE Then
         strFormat = VSP_CalTable("<10|<2|>2;", mdblWid, alngX)
         Vsp.FontBold = True

         If Sum2 > 0 Or m_cData.EXT_DISCOUNT_AMOUNT > 0 Or m_cData.VAT_AMOUNT > 0 Or mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Or mcolParam("DOCUMENT_TYPE") <> S_RETURN_DOCTYPE Then
            If mcolParam("BRANCH_ADDRESS") <> "" Then
               strPrint = "สถานที่ส่งของ   " & m_cData.APAR_NAME & "" & m_cData.ADD_BRANCH_NAME & "|รวมเป็นเงิน|" & FormatNumber(Sum3) & ";"
            Else
               strPrint = "หมายเหตุ " & m_cData.NOTE & "|รวมเป็นเงิน|" & FormatNumber(Sum3) & ";"
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If Sum2 > 0 Then
            If mcolParam("BRANCH_ADDRESS") <> "" Then
               strPrint = mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX") & "|ส่วนลด|" & FormatNumber(Sum2) & ";"
            Else
               strPrint = "|ส่วนลด|" & FormatNumber(Sum2) & ";"
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
                  
         If m_cData.EXT_DISCOUNT_AMOUNT > 0 Then
            If mcolParam("BRANCH_ADDRESS") <> "" And Sum2 = 0 Then
               strPrint = mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX") & "|ส่วนลดพิเศษ|" & FormatNumber(m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            Else
               strPrint = "|ส่วนลดพิเศษ|" & FormatNumber(m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If (Sum2 > 0 Or m_cData.EXT_DISCOUNT_AMOUNT > 0) And m_cData.VAT_AMOUNT > 0 And mcolParam("DOCUMENT_TYPE") <> RETURN_DOCTYPE Then
            strPrint = "|รวมหักส่วนลด|" & FormatNumber(Sum3 - Sum2 - m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If m_cData.VAT_AMOUNT > 0 Then
            If mcolParam("BRANCH_ADDRESS") <> "" And Sum2 = 0 And m_cData.EXT_DISCOUNT_AMOUNT = 0 Then
               strPrint = mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX") & "|ภาษี " & m_cData.VAT_PERCENT & " %" & "|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            Else
               strPrint = "|ภาษี " & m_cData.VAT_PERCENT & " %" & "|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            End If
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         Temp1 = Sum3 - Sum2 - m_cData.EXT_DISCOUNT_AMOUNT + m_cData.VAT_AMOUNT
         
         If mcolParam("BRANCH_ADDRESS") <> "" And Sum2 = 0 And m_cData.EXT_DISCOUNT_AMOUNT = 0 And m_cData.VAT_AMOUNT = 0 Then
            strPrint = mcolParam("BRANCH_ADDRESS") & " " & mcolParam("BRANCH_ADDRESS_PHONE") & " " & mcolParam("BRANCH_ADDRESS_FAX") & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
            strPrint = "(" & ThaiBaht(Temp1) & ")|รวมสุทธิ|" & FormatNumber(Temp1) & ";"
         Else
            strPrint = "(" & ThaiBaht(Temp1) & ")|รวมสุทธิ|" & FormatNumber(Temp1) & ";"
         End If
         Call Vsp.AddTable(strFormat, "", strPrint)
      ElseIf mcolParam("DOCUMENT_TYPE") = RECEIPT2_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = BILLS_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_RECEIPT2_DOCTYPE Or mcolParam("DOCUMENT_TYPE") = S_BILLS_DOCTYPE Then
         strFormat = VSP_CalTable(FOOTER_RECEIPT2, mdblWid, alngX)
         Vsp.FontBold = True
         strPrint = "|รวม|" & FormatNumber(Sum3) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
         
         If TempDiscount > 0 Then
            strPrint = "|ส่วนลดรับ|" & FormatNumber(TempDiscount) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If m_cData.EXT_DISCOUNT_AMOUNT > 0 Then
            strPrint = "|ส่วนลดพิเศษ|" & FormatNumber(m_cData.EXT_DISCOUNT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
            
         If m_cData.SUBTRACT_AMOUNT > 0 Then
            strPrint = "|ส่วนหัก|" & FormatNumber(m_cData.SUBTRACT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If m_cData.ADDITION_AMOUNT > 0 Then
            strPrint = "|ส่วนเพิ่ม|" & FormatNumber(m_cData.ADDITION_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         If m_cData.VAT_AMOUNT > 0 Then
            strPrint = "|ภาษี|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         strPrint = "(" & ThaiBaht(Sum3 - TempDiscount - m_cData.EXT_DISCOUNT_AMOUNT - m_cData.SUBTRACT_AMOUNT + m_cData.ADDITION_AMOUNT + m_cData.VAT_AMOUNT) & ")|รวมเงินสุทธิ|" & FormatNumber(Sum3 - TempDiscount - m_cData.EXT_DISCOUNT_AMOUNT - m_cData.SUBTRACT_AMOUNT + m_cData.ADDITION_AMOUNT + m_cData.VAT_AMOUNT) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)

      ElseIf mcolParam("DOCUMENT_TYPE") = CN_DOCTYPE Then
         strFormat = VSP_CalTable(FOOTER_CNDN, mdblWid, alngX)
         Vsp.FontBold = True
         strPrint = "|รวมลดหนี้|" & FormatNumber(Sum3) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
         
         If m_cData.VAT_AMOUNT > 0 Then
            strPrint = "|ภาษี|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         strPrint = "(" & ThaiBaht(Sum3 + m_cData.VAT_AMOUNT) & ")|รวมลดหนี้สุทธิ|" & FormatNumber(Sum3 + m_cData.VAT_AMOUNT) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
         
      ElseIf mcolParam("DOCUMENT_TYPE") = DN_DOCTYPE Then
         strFormat = VSP_CalTable(FOOTER_CNDN, mdblWid, alngX)
         Vsp.FontBold = True
         strPrint = "(" & ThaiBaht(Sum3) & ")|รวมเพิ่มหนี้|" & FormatNumber(Sum3) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
         
         If m_cData.VAT_AMOUNT > 0 Then
            strPrint = "|ภาษี|" & FormatNumber(m_cData.VAT_AMOUNT) & ";"
            Call Vsp.AddTable(strFormat, "", strPrint)
         End If
         
         strPrint = "(" & ThaiBaht(Sum3 + m_cData.VAT_AMOUNT) & ")|รวมลดหนี้สุทธิ|" & FormatNumber(Sum3 + m_cData.VAT_AMOUNT) & ";"
         Call Vsp.AddTable(strFormat, "", strPrint)
      End If

      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.PageWidth - Vsp.MarginRight, Y1, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
      
      Vsp.FontBold = False
      Call PrintBillFooter(alngX)
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   Rs.Close
   Set Rs = Nothing
   Set m_cData = Nothing
   
   genDoc = True
   
   If mcolParam("PREVIEW_TYPE") = 2 Then
      Vsp.PrintDoc (True)
   End If
   
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.Description
'   ErrorObj.RoutineName = RName
'   ErrorObj.ModuleName = MODULE_NAME
'   ErrorObj.LocalErrorMsg = "Run time error."
'   ErrorObj.SystemErrorMsg = mstrErrMsg
'   Call ErrorObj.ShowErrorLog(LOG_TO_FILE)
   Set Rs = Nothing
End Function

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
   If AllEditEndPage Then
      X2 = Vsp.PageWidth - Vsp.MarginRight
      Y2 = Vsp.CurrentY
         
      Call Vsp.DrawLine(Vsp.MarginLeft, Y1, Vsp.MarginLeft, Y2)
      Call Vsp.DrawLine(X1, Y1, X2, Y2)
      Call Vsp.DrawLine(Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY)
   End If
   
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub
Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdt As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long
   
   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdt = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
'   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   'Vsp.Paragraph = MODULE_DESC
   'Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdt
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub
