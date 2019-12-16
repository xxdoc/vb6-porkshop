VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuiConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GuiConfigs As Collection                                                                    '1

Public DOAddMenuItems As Collection                                                          '4
Public POAddMenuItems As Collection                                                          '5

Public ReceiptAddMenuItems As Collection                                                    '6
Public ReportMenuItems As Collection                                                            '7
Public QuoatationPrintMenuItems As Collection                                          '8
Public POPrintMenuItems As Collection                                                        '9
Public DOPrintMenuItems As Collection                                                        '10

Public IVPrintMenuItems As Collection                                                        '11
Public RCPrintMenuItems As Collection                                                        '12
Public RCPrintMenuItems2 As Collection                                                        '12
Public RCPackPrintMenuItems As Collection                                                        '12
Public RTPrintMenuItems As Collection                                           '13
Public CdnDbnPrintMenuItems As Collection                                              '14
Public BSPrintMenuItems As Collection                                                        '15
Public RoPrintMenuItems As Collection                                                        '15

Public PkglistPrintMenuItems As Collection                                                '16
Public MasterCommandMenuItems As Collection                                      '17
Public MainCommandMenuItems As Collection                                        '18
Public StockCommandMenuItems As Collection                                       '19
Public ProdCommandMenuItems As Collection                                          '20

Public LedgerCommandMenuItems As Collection                                       '21
Public GLCommandMenuItems As Collection                                             '22
Public CMCommandMenuItems As Collection                                             '23
Public PackageCommandMenuItems As Collection                                     '24
Public CostCommandMenuItems As Collection                                     '24

Public AdminCommandMenuItems As Collection                                       '26
Public TagetCommandMenuItems As Collection                                     '27
Public SellSubMenuItems As Collection                                                                    '28
Public PoSubMenuItems As Collection
Public SellReturnMenuItems As Collection                                                                    '29
Public CashMenuItems As Collection                                                                                 '30

Private Sub Class_Initialize()
   Set m_GuiConfigs = New Collection
   Set ReportMenuItems = New Collection
   Set DOAddMenuItems = New Collection
   
   Set POAddMenuItems = New Collection
   Set ReceiptAddMenuItems = New Collection
   Set QuoatationPrintMenuItems = New Collection
   Set POPrintMenuItems = New Collection
   Set DOPrintMenuItems = New Collection
   
   Set IVPrintMenuItems = New Collection
   Set RCPrintMenuItems = New Collection
   Set RCPrintMenuItems2 = New Collection
   Set RCPackPrintMenuItems = New Collection
   Set RTPrintMenuItems = New Collection
   Set CdnDbnPrintMenuItems = New Collection
   Set BSPrintMenuItems = New Collection
   Set RoPrintMenuItems = New Collection
   
   Set PkglistPrintMenuItems = New Collection
   Set CMCommandMenuItems = New Collection
   Set MasterCommandMenuItems = New Collection
   Set MainCommandMenuItems = New Collection
   Set StockCommandMenuItems = New Collection
   
   Set ProdCommandMenuItems = New Collection
   Set LedgerCommandMenuItems = New Collection
   Set PackageCommandMenuItems = New Collection
   Set AdminCommandMenuItems = New Collection
   Set CostCommandMenuItems = New Collection
   
   Set GLCommandMenuItems = New Collection
   Set TagetCommandMenuItems = New Collection
   Set SellSubMenuItems = New Collection
   Set PoSubMenuItems = New Collection
   Set SellReturnMenuItems = New Collection
   Set CashMenuItems = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_GuiConfigs = Nothing
   
   Set ReportMenuItems = Nothing
   Set DOAddMenuItems = Nothing
   Set POAddMenuItems = Nothing
   Set ReceiptAddMenuItems = Nothing
   Set QuoatationPrintMenuItems = Nothing
   Set POPrintMenuItems = Nothing
   Set DOPrintMenuItems = Nothing
   Set IVPrintMenuItems = Nothing
   Set RCPrintMenuItems = Nothing
   Set RCPrintMenuItems2 = Nothing
   Set RCPackPrintMenuItems = Nothing
   Set CdnDbnPrintMenuItems = Nothing
   Set BSPrintMenuItems = Nothing
   Set RoPrintMenuItems = Nothing
   
   Set CMCommandMenuItems = Nothing
   Set MasterCommandMenuItems = Nothing
   Set MainCommandMenuItems = Nothing
   Set StockCommandMenuItems = Nothing
   Set ProdCommandMenuItems = Nothing
   Set LedgerCommandMenuItems = Nothing
   Set PackageCommandMenuItems = Nothing
   Set AdminCommandMenuItems = Nothing
   Set CostCommandMenuItems = Nothing
   
   Set RTPrintMenuItems = Nothing
   Set PkglistPrintMenuItems = Nothing
   Set GLCommandMenuItems = Nothing
   Set TagetCommandMenuItems = Nothing
   Set SellSubMenuItems = Nothing
   Set PoSubMenuItems = Nothing
   Set SellReturnMenuItems = Nothing
   Set CashMenuItems = Nothing
End Sub

Private Sub CreateGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = New CGuiConfig
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
   
   Call m_GuiConfigs.add(TempC, Key)
   Set TempC = Nothing
End Sub

Private Sub SetGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(Key)
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
End Sub

Private Sub GenerateDOAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการใหม่ (ทีละรายการ)")
   TempID(1) = 1
   TempKey(2) = MapText("เพิ่มรายการใหม่ (เป็นชุด)")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   TempKey(4) = MapText("เลือกจากใบ PO")
   TempID(4) = 4
   TempKey(5) = MapText("เลือกจากใบเสนอราคา")
   TempID(5) = 5
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'lMenuChosen = oMenu.Popup("พิมพ์บนกระดาษเปล่า", "ปรับค่าหน้ากระดาษ", "-", "พิมพ์บนฟอร์มเปล่า (เหมือนจริง)", "พิมพ์บนฟอร์มเปล่า (ข้อมูล)", "ปรับค่าหน้ากระดาษ")
Private Sub GeneratePrintMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(50) As String
Dim TempID(50) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("พิมพ์กระดาษเปล่า")
   TempID(1) = 1
   TempKey(2) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(2) = 2
   TempKey(3) = MapText("-")
   TempID(3) = 3
   
   TempKey(4) = MapText("ใบเสร็จเป็นชุด (Preview)")
   TempID(4) = 4
   TempKey(5) = MapText("ใบเสร็จเป็นชุด (พิมพ์ทันที)")
   TempID(5) = 5
   TempKey(6) = MapText("ปรับค่าหน้ากระดาษ")
   TempID(6) = 6
   
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'oMenu.Popup("เพิ่มรายการใหม่ (ทีละรายการ)", "เพิ่มรายการใหม่ (เป็นชุด)")
Private Sub GeneratePOAddMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("เพิ่มรายการใหม่ (ทีละรายการ)")
   TempID(1) = 1
   TempKey(2) = MapText("เพิ่มรายการใหม่ (เป็นชุด)")
   TempID(2) = 2
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub

'oMenu.Popup("สร้างรายการใหม่", "-", "สร้างจากใบแจ้งหนี้", "-", "สร้างจากใบกำกับภาษี")
Private Sub GenerateAddReceiptMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem

   TempKey(1) = MapText("สร้างรายการใหม่ (ขายสด)")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("สร้างจากใบส่งของ (รับชำระหนี้)")
   TempID(3) = 3
   TempKey(4) = MapText("-")
   TempID(4) = 4
   TempKey(5) = MapText("สร้างจากใบกำกับภาษี")
   TempID(5) = 5
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         Set D = Nothing
      End If
   Next I
End Sub
Private Sub GenerateReportMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long

   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย")
   D.KEYWORD = ROOT_TREE & " A-1"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/ลูกค้า (SL001)")
   D.KEYWORD = ROOT_TREE & " 5-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '3
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการขายรายวันตามลูกค้า (SL002)")
   D.KEYWORD = ROOT_TREE & " 5-2"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '4
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการขายรายวัน (SL003)")
   D.KEYWORD = ROOT_TREE & " 5-3"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '5
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปการขายรายวัน (SL004)")
   D.KEYWORD = ROOT_TREE & " 5-4"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '6
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปขายตามสินค้า/พนักงาน (SL005)")
   D.KEYWORD = ROOT_TREE & " 5-5"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '7
'   Set D = New CMenuItem
'   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ/ลูกค้า (SL006)")
'   D.KEYWORD = ROOT_TREE & " 5-6"
'   D.PARENT_KEY = ROOT_TREE & " A-1"
'   D.ICON_INDEX1 = 1
'   D.ICON_INDEX2 = 2
'   Call Col.add(D)
'   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานค่าขนส่งตามทะเบียนรถ (SL006)")
   D.KEYWORD = ROOT_TREE & " A-1-1"
   D.PARENT_KEY = ROOT_TREE & " A-1"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '===
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบลูกหนี้")
   D.KEYWORD = ROOT_TREE & " A-2"
   D.PARENT_KEY = ROOT_TREE
   D.ICON_INDEX1 = 3
   D.ICON_INDEX2 = 3
   Call Col.add(D)
   Set D = Nothing
   
   '2
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานลูกหนี้ค้างชำระรายตัว (AR001)")
   D.KEYWORD = ROOT_TREE & " A-2-1"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานสรุปลูกหนี้ค้างชำระรายตัว (AR002)")
   D.KEYWORD = ROOT_TREE & " A-2-2"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการ์ดลูกหนี้ (AR003)")
   D.KEYWORD = ROOT_TREE & " A-2-3"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานวิเคราะห์อายุลูกหนี้ (AR004)")
   D.KEYWORD = ROOT_TREE & " A-2-4"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานการนำฝากธนาคาร (AR005)")
   D.KEYWORD = ROOT_TREE & " A-2-5"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเงินสดคงเหลือ (AR006)")
   D.KEYWORD = ROOT_TREE & " A-2-6"
   D.PARENT_KEY = ROOT_TREE & " A-2"
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing

   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateMasterCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักส่วนกลาง")
   D.KEYWORD = "MASTER_MAIN"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบบริหารบัญชี")
   D.KEYWORD = "MASTER_LEDGER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลหลักระบบคลัง")
   D.KEYWORD = "MASTER_INVENTORY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลหลัก")
   D.KEYWORD = "MASTER_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateMainCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
         
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลบริษัท")
   D.KEYWORD = "MAIN_ENTERPRISE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลลูกค้า")
   D.KEYWORD = "MAIN_CUSTOMER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลซัพพลายเออร์")
   D.KEYWORD = "MAIN_SUPPLIER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลพนักงาน")
   D.KEYWORD = "MAIN_EMPLOYEE"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลกลาง")
   D.KEYWORD = "MAIN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateProdCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("สูตรการผลิต")
   D.KEYWORD = "PRODUCT_FORMULA"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการผลิต")
   D.KEYWORD = "PRODUCT_JOB"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบการผลิต")
   D.KEYWORD = "PRODUCT_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
         
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateStockCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลรหัสคลัง")
   D.KEYWORD = "INVENTORY_PART"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลเอกสารคลัง")
   D.KEYWORD = "INVENTORY_DOC"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบคลัง")
   D.KEYWORD = "INVENTORY_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateAdminCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลกลุ่มผู้ใช้งาน")
   D.KEYWORD = "ADMIN_GROUP"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลผู้ใช้งาน")
   D.KEYWORD = "ADMIN_USER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานข้อมูลผู้ใช้งาน")
   D.KEYWORD = "ADMIN_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub

Private Sub GenerateLedgerCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานขาย")
   D.KEYWORD = "LEDGER_SELL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ระบบงานซื้อ")
   D.KEYWORD = "LEDGER_BUY"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานระบบบัญชี")
   D.KEYWORD = "LEDGER_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("โปรแกรมบัญชีการเงินอื่นๆ")
   D.KEYWORD = "LEDGER_PROGRAM"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   '===
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GeneratePackageCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลการตั้งราคาสินค้า")
   D.KEYWORD = "PACKAGE_DATA"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateTagetCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ข้อมูลเป้าการขายลูกค้า")
   D.KEYWORD = "TAGET_CUSTOMER"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
      
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("รายงานเป้าการขาย")
   D.KEYWORD = "TAGET_REPORT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Private Sub GenerateCashMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(20) As String
Dim TempID(20) As Long
Dim D As CMenuItem
   
   TempKey(1) = CashDocTypeToText(CHEQUE_REV)
   TempID(1) = CHEQUE_REV
   TempKey(2) = MapText("-")
   TempID(2) = -1
   
   TempKey(3) = CashDocTypeToText(CHEQUE_PAY)
   TempID(3) = CHEQUE_PAY
   TempKey(4) = MapText("-")
   TempID(4) = -1
   
   TempKey(5) = CashDocTypeToText(CASH_DEPOSIT)
   TempID(5) = CASH_DEPOSIT
   TempKey(6) = MapText("-")
   TempID(6) = -1
   
   TempKey(7) = CashDocTypeToText(POST_CHEQUE)
   TempID(7) = POST_CHEQUE
   TempKey(8) = MapText("-")
   TempID(8) = -1
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.add(D)
         
         Set D = Nothing
      End If
   Next I
End Sub
Private Sub GenerateSellSubExMenuItems(MenuMask As String, Col As Collection, Optional AddPoItem As Boolean = False)
Dim Mask As String
Dim I As Long
Dim TempKey(30) As String
Dim TempID(30) As Long
Dim D As CMenuItem
Dim Mr As CMasterRef
Dim TempColl As Collection

   Set TempColl = New Collection
   
   Call LoadMaster(Nothing, TempColl, , , MASTER_INVOICE_SUB)
   
   I = 0
   
   For Each Mr In TempColl
      I = I + 1
      TempKey((2 * I) - 1) = Mr.KEY_NAME
      TempID((2 * I) - 1) = Mr.KEY_ID
      
      If I <> TempColl.Count Then
         TempKey(2 * I) = MapText("-")
         TempID(2 * I) = 0
      End If
   Next Mr
   
   Set Mr = Nothing
   
   If AddPoItem Then
      Set D = New CMenuItem
      D.KEYWORD = "ขายสด"
      D.KEY_ID = 1
      Call Col.add(D)
      Set D = Nothing
      
      Set D = New CMenuItem
      D.KEYWORD = "-"
      D.KEY_ID = 0
      Call Col.add(D)
      Set D = Nothing
   End If
   
   For I = 1 To ((2 * TempColl.Count) - 1)
      Set D = New CMenuItem
      D.KEYWORD = TempKey(I)
      D.KEY_ID = TempID(I)
      Call Col.add(D)
      Set D = Nothing
   Next I
   
   Set TempColl = Nothing
End Sub
Private Sub GenerateSellReturnExMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(30) As String
Dim TempID(30) As Long
Dim D As CMenuItem
Dim TempColl As Collection
Dim Mr As CMasterRef
   
   Set TempColl = New Collection
   
   Call LoadMaster(Nothing, TempColl, , , MASTER_INVOICE_RETURN)
   
   I = 0
   
   For Each Mr In TempColl
      I = I + 1
      TempKey((2 * I) - 1) = Mr.KEY_NAME
      TempID((2 * I) - 1) = Mr.KEY_ID
      
      If I <> TempColl.Count Then
         TempKey(2 * I) = MapText("-")
         TempID(2 * I) = 0
      End If
   Next Mr
   
   Set Mr = Nothing
   
   For I = 1 To ((2 * TempColl.Count) - 1)
      Set D = New CMenuItem
      D.KEYWORD = TempKey(I)
      D.KEY_ID = TempID(I)
      Call Col.add(D)
      Set D = Nothing
   Next I
   
   Set TempColl = Nothing
End Sub
Private Sub GenerateCostCommandMenuItems(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim D As CMenuItem
Dim TempCount As Long
   
   '1
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ปรับต้นทุนมาตรฐาน")
   D.KEYWORD = "COST_STD"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("คำนวณต้นทุนคงเหลือและต้นทุนขาย")
   D.KEYWORD = "COST_CAPITAL"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   Set D = New CMenuItem
   D.MENU_TEXT = MapText("ปรับยอด STOCK แบบเป็นชุด")
   D.KEYWORD = "COST_STOCK-AMOUNT"
   D.PARENT_KEY = ""
   D.ICON_INDEX1 = 1
   D.ICON_INDEX2 = 2
   Call Col.add(D)
   Set D = Nothing
   
   '====
   TempCount = Col.Count
   While TempCount > 0
      If Mid(MenuMask, TempCount, 1) = "N" Then
         Call Col.Remove(TempCount)
      End If
      
      TempCount = TempCount - 1
   Wend
End Sub
Public Sub CreateGuiConfig(Programowner As String)
    
   '====
    Call CreateGuiConfigObj("LOGO_VIEW", False)
    Call CreateGuiConfigObj("AGENT_SERVER_CONNECT", False)
   
    Call SetGuiConfigObj("AGENT_SERVER_CONNECT", True)
   
   '====
   Call CreateGuiConfigObj("GL_VIEW", True)
   Call CreateGuiConfigObj("HR_VIEW", True)
   Call CreateGuiConfigObj("PRODUCTION_VIEW", True)
   Call CreateGuiConfigObj("PACKAGE_VIEW", True)
   Call CreateGuiConfigObj("LEDGER_VIEW", True)
   Call CreateGuiConfigObj("INVENTORY_VIEW", True)
   Call CreateGuiConfigObj("ENTERPRISE_INFO_EDIT", True)
   Call CreateGuiConfigObj("TAGET_VIEW", True)
   Call CreateGuiConfigObj("COST_VIEW", True)
   
   '====
   Call SetGuiConfigObj("HR_VIEW", False)
   Call SetGuiConfigObj("LEDGER_VIEW", True)
   Call SetGuiConfigObj("PACKAGE_VIEW", True)
   Call SetGuiConfigObj("TAGET_VIEW", True)
   
    Call SetGuiConfigObj("GL_VIEW", False)
   
   Call GenerateCashMenuItems("YYYYYYY", CashMenuItems, Programowner)
   
   Call GenerateSellSubExMenuItems("Y", SellSubMenuItems)
   Call GenerateSellSubExMenuItems("Y", PoSubMenuItems, True)
   Call GenerateSellReturnExMenuItems("Y", SellReturnMenuItems, Programowner)
   
   Call GenerateReportMenuItems("YYYYYYYYYYYYYYYYYY", ReportMenuItems, Programowner)
   Call GenerateDOAddMenuItems("YYNNN", DOAddMenuItems, Programowner)
   Call GeneratePOAddMenuItems("YY", POAddMenuItems, Programowner)
   Call GenerateAddReceiptMenuItems("YYYNN", ReceiptAddMenuItems, Programowner)
   
   Call GeneratePrintMenuItems("YY", DOPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("YY", RCPrintMenuItems, Programowner)
   Call GeneratePrintMenuItems("YY", RTPrintMenuItems, Programowner)
   
   Call GeneratePrintMenuItems("NNNYYY", RCPackPrintMenuItems, Programowner)
   
   
'   Call GeneratePrintMenuItems("YYNNNN", QuoatationPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("YYYNNNNYYYYYNNNNNNNNNNNNNYYYYYYY", POPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("YYNNNN", IVPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("YYYYYYNNNNNNNNNNNN", RCPrintMenuItems2, Programowner)
'   Call GeneratePrintMenuItems("YYYYYYNNNNNNNNYYYY", RCPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("NNNNNNNNNNNNNNNYYYYYYY", RCPackPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("YYNNNNNNNNNN", RTPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("YYNNNN", CdnDbnPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("YYYNNNNNNNNNNNNNNNNNNNNYY", BSPrintMenuItems, Programowner)
'   Call GeneratePrintMenuItems("NNNNNNNNNNNNYY", RoPrintMenuItems, Programowner)  'พิมพ์ใบรับสินค้า
   
   
   Call GenerateMainCommandMenuItems("YYYYY", MainCommandMenuItems, Programowner)
   
   Call GenerateMasterCommandMenuItems("YYYY", MasterCommandMenuItems, Programowner)
   
   Call GenerateAdminCommandMenuItems("YYN", AdminCommandMenuItems, Programowner)
   Call GenerateProdCommandMenuItems("YYY", ProdCommandMenuItems, Programowner)
   
   Call GenerateLedgerCommandMenuItems("YYY", LedgerCommandMenuItems, Programowner)
   
   Call GenerateStockCommandMenuItems("YYYYYY", StockCommandMenuItems, Programowner)
   Call GeneratePackageCommandMenuItems("Y", PackageCommandMenuItems, Programowner)
   Call GenerateCostCommandMenuItems("YYY", CostCommandMenuItems, Programowner)
   Call GenerateTagetCommandMenuItems("Y", TagetCommandMenuItems, Programowner)
End Sub
Public Function VerifyGuiConfig(KeyCode As String, Optional PrintError As Boolean = False) As Boolean
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(KeyCode)
   VerifyGuiConfig = TempC.ENABLE_FLAG
   
   If Not TempC.ENABLE_FLAG Then
      If PrintError Then
         glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
         glbErrorLog.ShowUserError
      End If
   End If
End Function
Public Function ShowWindowCaption() As String
   ShowWindowCaption = MapText("ระบบบริหารงาน PorkShop") & " " & Command & " " & glbParameterObj.DBFile
End Function
