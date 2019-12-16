Attribute VB_Name = "modMain"
Option Explicit

Public Enum PAYMENT_TYPE
   CASH_PMT = 1                                    ' เงินสด ทดสอบการเปลี่ยน
   CHEQUE_HAND_PMT = 2                                'เช็คเข้ามือก่อน
   CHEQUE_BANK_PMT = 3                                'เช็คเข้าธนาคารโดยตรง
   BANKTRF_PMT = 4                           'โอนเงิน
End Enum

Public Enum CASH_DOC_TYPE
   CHEQUE_REV = 1                                                          'เช็ครับ
   CHEQUE_PAY = 2                                                          'เช็คจ่าย
   CASH_DEPOSIT = 3                                                     'ใบนำฝาก
   POST_CHEQUE = 4                                                      'ใบ Post Cheque      เช็คที่เข้าธนาคารเมื่อได้รับเงินแล้วแสดงว่าได้ทำการ POST แล้ว
End Enum

Public Enum POST_TYPE
   POST_CLEAR = 1                                                      ' ใบเครียร์ ของ เช็คจากลูกค้า เมื่อได้เงินแล้วจริงๆ ในธนาคารจะถือว่า เครียร์
   WAITING_CLEAR = 2                                                      'เช็ครอเรียกเก็บ
   PASSED_CLEAR = 3                                                      'เช็คผ่านแล้ว
End Enum

Public Enum FIELD_TYPE
   INT_TYPE = 1
   MONEY_TYPE = 2
   DATE_TYPE = 3
   STRING_TYPE = 4
   BOOLEAN_TYPE = 5
End Enum

Public Enum TAGET_TYPE
   TAGET_CUSTOMER = 2
End Enum

Public Enum FIELD_CAT
   ID_CAT = 1
   MODIFY_DATE_CAT = 2
   CREATE_DATE_CAT = 3
   MODIFY_BY_CAT = 4
   CREATE_BY_CAT = 5
   DATA_CAT = 6
   TEMP_CAT = 7
End Enum

Public Enum SHOW_MODE_TYPE
   SHOW_ADD = 1
   SHOW_EDIT = 2
   SHOW_VIEW = 3
   SHOW_VIEW_ONLY = 4
End Enum

Public Enum TEXT_BOX_TYPE
   TEXT_STRING = 1
   TEXT_INTEGER = 2
   TEXT_FLOAT = 3
   TEXT_FLOAT_MONEY = 4
   TEXT_INTEGER_MONEY = 5
End Enum

Public Enum MASTER_TYPE
   MASTER_COUNTRY = 1
   MASTER_SEX = 2
   MASTER_CUSTYPE = 3
   MASTER_SUPTYPE = 5
   MASTER_DEPARTMENT = 10
   MASTER_UNIT = 11
   MASTER_STOCKTYPE = 12
   MASTER_STOCKGROUP = 13
   MASTER_LOCATION = 14
   MASTER_BANK = 16
   MASTER_BBRANCH = 17
   MASTER_CHEQUE_TYPE = 18
   MASTER_CNDN_REASON = 19
   MASTER_INVOICE_SUB = 23
   MASTER_INVOICE_RETURN = 24
   MASTER_SUBTRACT = 25
   MASTER_BANK_ACCOUNT = 26
   MASTER_BACCOUNT_TYPE = 27
   MASTER_ADDITION = 28
   MASTER_CUSGROUP = 32
   MASTER_INVENTORY_SUB_TYPE = 34
End Enum

Public Enum SELL_BILLING_DOCTYPE                   'DocumentTypeขายเป็นตามด้านล่าง ส่วนฝั่งซื้อจะ + 100
   QUOATATION_DOCTYPE = 1
   PO_DOCTYPE = 2
   INVOICE_DOCTYPE = 3
   RECEIPT1_DOCTYPE = 4
   RECEIPT2_DOCTYPE = 5
   RETURN_DOCTYPE = 6
   CN_DOCTYPE = 7
   DN_DOCTYPE = 8
   BILLS_DOCTYPE = 9
   RECEIPT3_DOCTYPE = 10
   
   S_QUOATATION_DOCTYPE = 101
   S_PO_DOCTYPE = 102
   S_INVOICE_DOCTYPE = 103
   S_RECEIPT1_DOCTYPE = 104
   S_RECEIPT2_DOCTYPE = 105
   S_RETURN_DOCTYPE = 106
   S_CN_DOCTYPE = 107
   S_DN_DOCTYPE = 108
   S_BILLS_DOCTYPE = 109
   
End Enum

Public Enum INVENTORY_DOCTYPE
   IMPORT_DOCTYPE = 1
   EXPORT_DOCTYPE = 2
   TRANSFER_DOCTYPE = 3
   ADJUST_DOCTYPE = 4
End Enum


Public Enum UNIQUE_TYPE
   PACKAGE_NO = 1
   PACKAGE_DESC = 2
   PACKAGE_MASTER_FLAG = 3
   DOCUMENT_NO_UNIQUE = 6
   APARCODE_UNIQUE = 7
   INVENTORY_DOC_NO = 8
   PARTNO_UNIQUE = 9
   MASTER_CODE = 11
   MASTER_NAME = 12
   EMPCODE_UNIQUE = 17
   MASTER_VALID_NO = 20
   FORMULA_NO = 21
   JOB_NO = 22
End Enum

Public Enum DEALER_TYPE_AREA
   SILVER = 10
   SILVER_PLUS = 15
   SILVER_PLUS_PLUS = 20
   GOLD_MUNUS = 25
   GOLD = 30
   GOLD_PLUS = 35
   GOLD_PLUS_PLUS = 40
   PLATINUM_MUNUS = 45
   PLATINUM = 50
   HEADER_GROUP = 100
End Enum

Public GLB_GRID_COLOR As Long
Public GLB_NORMAL_COLOR As Long
Public GLB_ALERT_COLOR As Long
Public GLB_SHOW_COLOR As Long
Public GLB_FORM_COLOR As Long
Public GLB_HEAD_COLOR As Long
Public GLB_GRIDHD_COLOR As Long
Public GLB_MANDATORY_COLOR As Long

Public glbErrorLog As clsErrorLog
Public glbDatabaseMngr As clsDatabaseMngr
Public glbParameterObj As clsParameter
Public glbUser As CUser
Public glbEnterpriseID As Long
Public glbGuiConfigs As CGuiConfigs
Public glbEnterPrise As CEnterprise
Public glbDaily As clsDaily
Public glbSetting As clsGlobalSetting
Public glbAccessRight As Collection


Public LoadPackageColl As Collection
Public m_CustomerColl As Collection
Public m_SupplierColl As Collection
Public m_EmployeeColl As Collection
Public m_LocationColl As Collection
Public InventorySubTypecoll As Collection

Public glbLockDate As CLockDate

'Public Const GLB_FONT = "JasmineUPC"

Public Const GLB_FONT = "JasmineUPC"
Public Const GLB_FONT_EX = "Cordia New"
Public Const ROOT_TREE = "Root"
Public Const DUMMY_KEY = 27
Public Const PROJECT_NAME = "Exclusive System Software"

'===================== For clear treeview =========================
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd _
    As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const TV_FIRST As Long = &H1100
Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Const TVGN_ROOT As Long = &H0
Const WM_SETREDRAW As Long = &HB
'===================== For clear treeview =========================

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ChangeQuote(StrQ As String) As String
   ChangeQuote = Replace(StrQ, "'", "''")
End Function

Public Function DateToStringInt(D As Date) As String
   If D = -1 Then
      DateToStringInt = "9999-99-99 99:99:99"
   ElseIf D = -2 Then
      DateToStringInt = "0000-00-00 00:00:00"
   Else
      DateToStringInt = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   End If
End Function
Public Function GenerateInsertSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String

   SQL = "INSERT INTO " & O.TableName & vbCrLf & " (" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ") " & vbCrLf & "VALUES " & vbCrLf & "(" & vbCrLf
         Else
            Sep = ", " & vbCrLf
         End If
         
         SQL = SQL & Tf.FieldName & Sep
      End If
   Next Tf
   
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ")"
         Else
            Sep = ", " & vbCrLf
         End If
''debug.print "---" & Tf.FieldName
         SQL = SQL & Tf.TransformToSQLString & Sep
''debug.print "---" & Tf.GetValue
      End If
   Next Tf
   
   GenerateInsertSQL = SQL
End Function

Public Function GenerateUpdateSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String
Dim TempKeyName As String
Dim TempKeyVal As Long

   SQL = "UPDATE " & O.TableName & " SET" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT And Tf.FieldCat <> CREATE_DATE_CAT And Tf.FieldCat <> CREATE_BY_CAT Then
         If Tf.FieldCat = ID_CAT Then
            TempKeyName = Tf.FieldName
            TempKeyVal = Tf.GetValue
         Else
            If Tf.FieldCat = MODIFY_BY_CAT Then
               Sep = "" & vbCrLf
            Else
               Sep = ", " & vbCrLf
            End If
            
            SQL = SQL & Tf.FieldName & " = " & Tf.TransformToSQLString & Sep
         End If
      End If
   Next Tf
      
   SQL = SQL & "WHERE " & TempKeyName & " = " & TempKeyVal
   
   GenerateUpdateSQL = SQL
End Function
Public Sub PopulateInternalField(ShowMode As SHOW_MODE_TYPE, O As Object)
Dim Tf As CTableField
Dim TempID As Long
Dim InternalDate As String

   For Each Tf In O.m_FieldList
      If Tf.FieldCat = ID_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetSeqID(O.SequenceName, TempID, glbErrorLog)
            Call Tf.SetValue(TempID)
         End If
      ElseIf Tf.FieldCat = CREATE_DATE_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         End If
      ElseIf Tf.FieldCat = MODIFY_DATE_CAT Then
         'If ShowMode = SHOW_EDIT Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         'End If
      ElseIf Tf.FieldCat = CREATE_BY_CAT Then
         If ShowMode = SHOW_ADD Then
            Call Tf.SetValue(glbUser.USER_ID)
         End If
      ElseIf Tf.FieldCat = MODIFY_BY_CAT Then
         'If ShowMode = SHOW_EDIT Then
            Call Tf.SetValue(glbUser.USER_ID)
         'End If
      End If
   Next Tf
End Sub

Public Function NVLD(Value As Variant, I As Double) As Double
On Error Resume Next

   If IsNull(Value) Then
      NVLD = I
   Else
      NVLD = Value
   End If
End Function

Public Function NVLS(Value As Variant, S As String) As String
On Error Resume Next

   If IsNull(Value) Then
      NVLS = S
   Else
      NVLS = Value
   End If
End Function

Public Function NVLI(Value As Variant, I As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLI = I
   Else
      NVLI = Value
   End If
End Function

Public Function EnableForm(Frm As Form, En As Boolean)
   If Frm Is Nothing Then
      Exit Function
   End If
   
   Frm.Enabled = En
   If En Then
      Screen.MousePointer = vbArrow
   Else
      Frm.Refresh
      DoEvents
      Screen.MousePointer = 11
   End If
End Function

Public Function CryptString(strInput As String, strKey As String, action As Boolean)
Dim I As Integer, C As Integer
Dim strOutput As String

If Len(strKey) Then
    For I = 1 To Len(strInput)
        C = Asc(Mid$(strInput, I, 1))
        If action Then
            C = C + Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        Else: C = C - Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        End If
        strOutput = strOutput & Chr$(C And &HFF)
    Next I
Else
    strOutput = strInput
End If
CryptString = strOutput
End Function

Public Function EncryptText(PText As String) As String
   EncryptText = CryptString(PText, "GENETICOTHELLO", True)
End Function

Public Function DecryptText(CText As String) As String
   DecryptText = CryptString(CText, "GENETICOTHELLO", False)
End Function
Public Sub InitTextBox(T As TextBox, msg As String, Optional Password As String = "")
   T.PasswordChar = Password
   T.FontSize = 12
   T.FontName = "MS Sans Serif"
   T.Text = msg
   T.BackColor = GLB_GRID_COLOR
   'T.FontBold = True
End Sub
Public Sub InitDialogButton(B As CommandButton, Caption As String)
   B.Caption = Caption
   B.FontBold = True
   B.FontSize = 14
   B.FontName = GLB_FONT
   
   B.BackColor = &HFFFFFF
End Sub

Public Sub SetEnableDisableTextBox(T As TextBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
      T.FontBold = False
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
      T.FontBold = True
   End If
End Sub

Public Sub SetEnableDisableComboBox(T As ComboBox, En As Boolean)
   If En Then
      T.TabStop = True
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.TabStop = False
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableButton(B As SSCommand, En As Boolean)
   If En Then
      B.Enabled = True
      B.BackColor = GLB_GRID_COLOR
   Else
      B.Enabled = False
      B.BackColor = &H8000000F
   End If
End Sub

Public Function ConfirmExit(HasEdit As Boolean) As Boolean
   If Not HasEdit Then
      ConfirmExit = True
   Else
      glbErrorLog.LocalErrorMsg = "ท่านต้องการจะออก โดยไม่มีการบันทึกข้อมูลใช่หรือไม่"
      If glbErrorLog.AskMessage = vbYes Then
         ConfirmExit = True
      Else
         ConfirmExit = False
      End If
   End If
End Function
Public Function ConfirmSave() As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการการบันทึกข้อมูลใช่หรือไม่"
   If glbErrorLog.AskMessage = vbYes Then
      ConfirmSave = True
   Else
      ConfirmSave = False
   End If
End Function
Public Function ThaiBaht(ByVal pamt As Double) As String
Dim valstr As String, vLen As Integer, vno As Integer, syslge As String
Dim I As Integer, j As Integer, V As Integer
Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
Dim vword(20) As String

 If pamt <= 0# Then
   ThaiBaht = ""
   Exit Function
 End If
 valstr = Trim(Format$(pamt, "##########0.00"))
 vLen = Len(valstr) - 3
 For I = 1 To 20
     vword(I) = ""
 Next I
wnumber(1) = "หนึ่ง": wnumber(2) = "สอง": wnumber(3) = "สาม": wnumber(4) = "สี่"
wnumber(5) = "ห้า": wnumber(6) = "หก": wnumber(7) = "เจ็ด": wnumber(8) = "แปด"
wnumber(9) = "เก้า": wdigit(1) = "บาท": wdigit(2) = "สิบ": wdigit(3) = "ร้อย": wdigit(4) = "พัน"
wdigit(5) = "หมื่น": wdigit(6) = "แสน": wdigit(7) = "ล้าน": spcdg(1) = "สตางค์": spcdg(2) = "เอ็ด"
spcdg(3) = "ยี่": spcdg(4) = "ถ้วน"
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
        vword(I) = ""
        If (vLen - I + 1) = 7 Then
            vword(I) = wdigit(7)             '--ล้าน
        End If
    Else
        If (vLen - I + 1) > 7 Then
            j = vLen - I - 5               '--เกินหลักล้าน
        Else
            j = vLen - I + 1               '--หลักแสน
        End If
        vword(I) = wnumber(vno) + wdigit(j) '-30ถึง90
        If vno = 1 And j = 2 Then
            vword(I) = wdigit(2)             '--สิบ
        End If
        If vno = 2 And j = 2 Then
            vword(I) = spcdg(3) + wdigit(j)  '--ยี่สิบ
        End If
        If j = 1 Then                       ' สิยเอ็ค -->เก้าสิบเอ็ด
            vword(I) = wnumber(vno)
            If vno = 1 And vLen > 1 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2)
                End If
            End If
        End If
        If j = 7 Then         '-แก้บักกรณี 11,111,111.00 สิบเอ็ด
            vword(I) = wnumber(vno) + wdigit(j)   '-ล้าน
            If vno = 1 And vLen > 7 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2) + wdigit(j)
                End If
            End If
        End If
    End If
Next I
    
If Int(pamt) > 0 Then
       vword(vLen) = vword(vLen) + wdigit(1)
End If
 '--------------ทศนิยม --------------
valstr = Mid$(valstr, vLen + 2, 2)
vLen = Len(valstr)
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
           vword(I + 10) = ""
    Else
           j = vLen - I + 1
           vword(I + 10) = wnumber(vno) + wdigit(j)
        If vno = 1 And j = 2 Then
              vword(I + 10) = wdigit(2)
        End If
        If vno = 2 And j = 2 Then
              vword(I + 10) = spcdg(3) + wdigit(j)
        End If
        If j = 1 Then
            If vno = 1 And Int(Val(Mid$(valstr, I - 1, 1))) <> 0 Then
                 vword(I + 10) = spcdg(2)
            Else
                 vword(I + 10) = wnumber(vno)
            End If
        End If
    End If
Next I
If pamt <> 0 Then
    If Val(valstr) = 0 Then
        vword(13) = spcdg(4)
    Else
        vword(13) = spcdg(1)
    End If
End If

 '*** เผื่อใช้กรณียาวมาก และต้องการตัดประโยค
 valstr = ""
 For I = 1 To 20
    'IF LEN(valstr) < 70 AND LEN(valstr + vword(i)) > 70 Then
    '   valstr = valstr + REPLICATE(" ",70 - LEN(valstr))
    'END IF
    valstr = valstr + vword(I)
 Next I
 'valstr='('+valstr+')'
 ThaiBaht = (valstr)
End Function

Public Function WildCard(WStr As String, SubLen As Long, NewStr As String) As Boolean
Dim Tmp As String
   Tmp = Trim(WStr)
   If Tmp = "" Then
      WildCard = False
      Exit Function
   End If
   
   If Mid(Tmp, Len(Tmp)) = "%" Then
      SubLen = Len(Tmp) - 1
      NewStr = Mid(Tmp, 1, SubLen)
      
      WildCard = True
   Else
      WildCard = False
   End If
End Function

Public Function FormatString(S As String, Patch As String, L As Long) As String
Dim Temp As String
Dim Start As Long
Dim I As Long
Dim j As Long

   Temp = Space(L)
   Call Replace(Temp, " ", Patch)
   j = 0
   Start = (L - Len(S)) \ 2
   
   For I = 1 To L
      If I < Start Then
         Mid(Temp, I) = Patch
      Else
         If I > Start + Len(S) Then
            Mid(Temp, I) = Patch
         Else
            j = j + 1
            Mid(Temp, I) = Mid(S, j)
         End If
      End If
   Next I
   
   FormatString = Temp
End Function

Public Function FormatNumber(n As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "0") As String
Dim T As Double
Dim TempStr As String
Dim I As Long

   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(n) Then
      T = 0
   Else
      T = n
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumber = ZeroString & TempStr
      Else
         FormatNumber = ZeroString
      End If
   ElseIf Quat Then
      FormatNumber = Format(T, "#,##0" & TempStr)
   Else
      FormatNumber = Format(T, "0" & TempStr)
   End If
End Function
Public Function FormatNumberToNull(n As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "") As String
Dim T As Double
Dim TempStr As String
Dim I As Long

   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(n) Then
      T = 0
   Else
      T = n
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumberToNull = ZeroString & TempStr
      Else
         FormatNumberToNull = ZeroString
      End If
   ElseIf Quat Then
      FormatNumberToNull = Format(T, "#,##0" & TempStr)
   Else
      FormatNumberToNull = Format(T, "0" & TempStr)
   End If
End Function
Public Function ReverseFormatNumber(n As String) As Double
   ReverseFormatNumber = Val(Replace(n, ",", ""))
End Function

Public Function IDToListIndex(Cbo As ComboBox, ID As Long) As Long
Dim I As Long
Dim Temp As String

   IDToListIndex = -1
   For I = 0 To Cbo.ListCount - 1
      If InStr(Cbo.ItemData(I), ":") <= 0 Then
         Temp = Cbo.ItemData(I)
      Else
         Temp = Mid(Cbo.ItemData(I), 1, InStr(Cbo.ItemData(I), ":") - 1)
      End If
      If Temp = ID Then
         IDToListIndex = I
      End If
   Next I
End Function

Public Function InternalDateToDate(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDate = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDate = -2
      Exit Function
   End If
   
   If Len(IntDate) < 19 Then
      InternalDateToDate = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = Mid(IntDate, 12, 2)
   MMStr = Mid(IntDate, 15, 2)
   SSStr = Mid(IntDate, 18, 2)
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDate = DateSerial(y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx2(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If Len(IntDate) < 10 Then
      InternalDateToDateEx2 = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
      
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
      
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx2 = DateSerial(y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If Len(IntDate) < 8 Then
      InternalDateToDateEx = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 5, 2)
   DStr = Mid(IntDate, 7, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   y = Val(YStr) - 543
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx = DateSerial(y, M, D) + TimeSerial(HH, MM, SS)
End Function
Public Function InternalDateToDateExGrid(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long
      
   If Len(IntDate) < 8 Then
      InternalDateToDateExGrid = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 7, 4)
   MStr = Mid(IntDate, 4, 2)
   DStr = Mid(IntDate, 1, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   y = Val(YStr) - 543
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateExGrid = DateSerial(y, M, D) + TimeSerial(HH, MM, SS)
End Function
Public Function ReFormatDate(DStr As String) As String
Dim YYYY As String
Dim MM As String
Dim dd As String

   YYYY = Mid(DStr, 5, 4)
   MM = Mid(DStr, 3, 2)
   dd = Mid(DStr, 1, 2)
   
   ReFormatDate = YYYY & MM & dd
End Function

Public Sub SetSelect(T As TextBox)
   T.SelStart = 0
   T.SelLength = Len(T.Text)
End Sub

Public Sub InitCombo(C As ComboBox)
   C.FontSize = 12
   C.FontName = "MS Sans Serif"
   C.BackColor = GLB_GRID_COLOR
End Sub
Public Function IntToThaiMonth(M As Long, Optional S As Long = -1) As String
   If glbParameterObj Is Nothing Then
      Exit Function
   End If
   
   If M = 1 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ม.ค."
         Else
            IntToThaiMonth = "มกราคม"
         End If
      Else
         IntToThaiMonth = "January"
      End If
   ElseIf M = 2 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ก.พ."
         Else
            IntToThaiMonth = "กุมภาพันธ์"
         End If
      Else
         IntToThaiMonth = "February"
      End If
      
   ElseIf M = 3 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "มี.ค."
         Else
            IntToThaiMonth = "มีนาคม"
         End If
      Else
         IntToThaiMonth = "March"
      End If
      
   ElseIf M = 4 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "เม.ย."
         Else
            IntToThaiMonth = "เมษายน"
         End If
      Else
         IntToThaiMonth = "April"
      End If
      
   ElseIf M = 5 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "พ.ค."
         Else
            IntToThaiMonth = "พฤษภาคม"
         End If
      Else
         IntToThaiMonth = "May"
      End If
      
   ElseIf M = 6 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "มิ.ย."
         Else
            IntToThaiMonth = "มิถุนายน"
         End If
      Else
         IntToThaiMonth = "June"
      End If
      
   ElseIf M = 7 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ก.ค."
         Else
            IntToThaiMonth = "กรกฎาคม"
         End If
      Else
         IntToThaiMonth = "July"
      End If
      
   ElseIf M = 8 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ส.ค."
         Else
            IntToThaiMonth = "สิงหาคม"
         End If
      Else
         IntToThaiMonth = "August"
      End If
      
   ElseIf M = 9 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ก.ย."
         Else
            IntToThaiMonth = "กันยายน"
         End If
      Else
         IntToThaiMonth = "September"
      End If
      
   ElseIf M = 10 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ต.ค."
         Else
            IntToThaiMonth = "ตุลาคม"
         End If
      Else
         IntToThaiMonth = "October"
      End If
      
   ElseIf M = 11 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "พ.ย."
         Else
            IntToThaiMonth = "พฤศจิกายน"
         End If
      Else
         IntToThaiMonth = "November"
      End If
      
   ElseIf M = 12 Then
      If glbParameterObj.Language = 1 Then
         If S = 1 Then
            IntToThaiMonth = "ธ.ค."
         Else
            IntToThaiMonth = "ธันวาคม"
         End If
      Else
         IntToThaiMonth = "December"
      End If
   Else
      IntToThaiMonth = ""
   End If
End Function

Public Function Minus2Zero(A As Double) As Long
   If A < 0 Then
      Minus2Zero = 0
   Else
      Minus2Zero = A
   End If
End Function

Public Function Zero2One(A As Double) As Long
   If A = 0 Then
      Zero2One = 1
   Else
      Zero2One = A
   End If
End Function

Public Sub ClearTreeView(ByVal tvHwnd As Long)
Dim lNodeHandle As Long

    'Turn off redrawing on the Treeview for more speed improvements
    SendMessageLong tvHwnd, WM_SETREDRAW, False, 0

    Do
        lNodeHandle = SendMessageLong(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
         If lNodeHandle > 0 Then
            SendMessageLong tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
         Else
            Exit Do
         End If
    Loop

    'Turn on redrawing on the Treeview
    SendMessageLong tvHwnd, WM_SETREDRAW, True, 0
End Sub

Public Function Minus2Flag(A As Double) As String
   If A < 0 Then
      Minus2Flag = "Y"
   Else
      Minus2Flag = "N"
   End If
End Function

Public Sub InitNormalLabel(L As Label, Caption As String, Optional Color As Long = 0)
   L.Caption = ""
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub

Public Sub SetTextLenType(T As TextBox, TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      T.Alignment = 1
   End If
   
   T.Tag = TT
   T.MaxLength = L
End Sub

Public Sub Main()
Dim TempDB As String
   GLB_GRID_COLOR = RGB(255, 255, 250)
   GLB_NORMAL_COLOR = RGB(0, 0, 0)
   GLB_ALERT_COLOR = RGB(255, 0, 0)
   GLB_FORM_COLOR = RGB(180, 200, 200)
   GLB_HEAD_COLOR = GLB_FORM_COLOR
   GLB_GRIDHD_COLOR = RGB(149, 194, 240)
   GLB_SHOW_COLOR = RGB(0, 0, 240)
   GLB_MANDATORY_COLOR = RGB(0, 0, 255)
   
   If App.PrevInstance = True Then
      glbErrorLog.LocalErrorMsg = "โปรแกรมเดิมได้ถูกรันก่อนหน้านี้แล้ว"
      glbErrorLog.ShowUserError

      Set glbErrorLog = Nothing
      Exit Sub
   End If
   
   Set glbErrorLog = New clsErrorLog
   Set glbDatabaseMngr = New clsDatabaseMngr
   Set glbUser = New CUser
   Set glbParameterObj = New clsParameter
   
   Set glbEnterPrise = New CEnterprise
   Set glbDaily = New clsDaily
   Set glbSetting = New clsGlobalSetting
   Set glbAccessRight = New Collection
   
   Set LoadPackageColl = New Collection
   Set m_CustomerColl = New Collection
   Set m_SupplierColl = New Collection
   Set m_LocationColl = New Collection
   Set m_EmployeeColl = New Collection
   Set InventorySubTypecoll = New Collection
   Set glbLockDate = New CLockDate
   
   glbEnterpriseID = 20
   
   TempDB = glbParameterObj.DBFile
   
   If Not (glbDatabaseMngr.ConnectDatabase(TempDB, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog)) Then
      frmDBSetting.UserName = glbParameterObj.UserName
      frmDBSetting.Password = glbParameterObj.Password
      frmDBSetting.FileDb = TempDB
      frmDBSetting.Header = " ไม่สามารถเชื่อต่อฐานข้อมูลได้ "

      Load frmDBSetting
      frmDBSetting.Show 1
      If frmDBSetting.OKClick Then
         glbParameterObj.UserName = frmDBSetting.UserName
         glbParameterObj.Password = frmDBSetting.Password
         
         glbParameterObj.DBFile = frmDBSetting.FileDb
      Else
         Unload frmDBSetting
         Set frmDBSetting = Nothing
         
         Unload frmSplash
         Set frmSplash = Nothing
         
         Call ReleaseAll
         End
      End If
      Unload frmDBSetting
      Set frmDBSetting = Nothing
   End If
   
   Set glbGuiConfigs = New CGuiConfigs
   Call glbGuiConfigs.CreateGuiConfig("")
   
   Load frmWinPricingMain
   frmWinPricingMain.Show
   
'   Unload frmWinPricingMain
'   Set frmWinPricingMain = Nothing
End Sub

Public Function MapText(msg As String) As String
   MapText = msg
End Function

Public Sub InitMainButton(B As SSCommand, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.Name = GLB_FONT
   B.Font3D = ssInsetLight
   B.BackColor = RGB(255, 255, 255)
   B.ButtonStyle = ssActiveBorders
   B.MousePointer = ssCustom
   B.MouseIcon = LoadPicture(glbParameterObj.ButtonCursor)
End Sub

Private Function GetParentKey(Acc As String, TopFlag As Boolean) As String
Dim I As Long
Dim j As Long

   For I = 1 To Len(Acc)
      If Mid(Acc, I, 1) = "_" Then
         j = I
      End If
   Next I
   
   If j > 1 Then
      GetParentKey = Mid(Acc, 1, j - 1)
      TopFlag = False
   Else
      GetParentKey = ""
      TopFlag = True
   End If
End Function
Private Sub GetParentItemDesc(Acc As String, Ri As CRightItem, ReportName As String)
   Call Ri.SetFieldValue("DEFAULT_VALUE", "N")
   
   If Acc = "ADMIN" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบผู้ใช้งาน")
   ElseIf Acc = "ADMIN_GROUP" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "กลุ่มข้อมูลผู้ใช้งาน")
   ElseIf Acc = "ADMIN_USER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลผู้ใช้งาน")
   
   
   ElseIf Acc = "MASTER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบข้อมูลหลัก")
   ElseIf Acc = "MASTER_MAIN" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลหลักส่วนกลาง")
   ElseIf Acc = "MASTER_MAIN_QUERY" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ค้นหาข้อมูลหลักส่วนกลาง")
   ElseIf Acc = "MASTER_LEDGER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลหลักบัญชี")
   ElseIf Acc = "MASTER_INVENTORY" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลหลักคลัง")
   ElseIf Acc = "MASTER_PRODUCTION" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลหลักการผลิต")
   
   ElseIf Acc = "MAIN" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลส่วนกลาง")
   ElseIf Acc = "MAIN_ENTERPRISE" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลบริษัท")
   ElseIf Acc = "MAIN_CUSTOMER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลลูกค้า")
   ElseIf Acc = "MAIN_SUPPLIER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลซัพพลายเออร์")
   ElseIf Acc = "MAIN_EMPLOYEE" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลพนักงาน")
   
   
   ElseIf Acc = "PRODUCT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "การผลิต")
   ElseIf Acc = "PRODUCT_FORMULA" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลสูตรการผลิต")
   ElseIf Acc = "PRODUCT_JOB" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลใบสั่งผลิต")
   
   ElseIf Acc = "LEDGER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบบัญชี")
   ElseIf Acc = "LEDGER_SELL" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบบัญชีขาย")
   ElseIf Acc = "LEDGER_BUY" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบบัญชีซื้อ")
   ElseIf Acc = "LEDGER_CASH" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบข้อมูลการเงิน")
   ElseIf Acc = "LEDGER_PROGRAM" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "โปรแกรมบัญชีการเงินอื่นๆ")
      
   ElseIf Acc = "INVENTORY" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบบริหารคลัง")
   ElseIf Acc = "INVENTORY_PART" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "รหัสคลัง(สินค้า/วัตถุดิบ)")
   ElseIf Acc = "INVENTORY_DOC" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "เอกสารคลัง")
   ElseIf Acc = "INVENTORY_BALANCE" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลการตรวจสอบข้อมูล")
   ElseIf Acc = "COMMISSION_TABLE" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ตารางค่าคอมมิตชั่น")
   ElseIf Acc = "COMMISSION_CHART" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "แผนภูมิการจัดคิดคอมมิตชั่น")
   ElseIf Acc = "PACKAGE" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบการตั้งราคาสินค้า")
   ElseIf Acc = "PACKAGE_DATA" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลการตั้งราคาสินค้า")
      
   ElseIf Acc = "TAGET" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "เป้าการขาย")
   ElseIf Acc = "TAGET_CUSTOMER" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ข้อมูลเป้าการขายลูกค้า")
   
   ElseIf Acc = "COST" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ระบบต้นทุน")
   ElseIf Acc = "COST_STD" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ปรับต้นทุนมาตรฐาน")
   ElseIf Acc = "COST_CAPITAL" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "คำนวณต้นทุนคงเหลือและต้นทุนขาย")
   ElseIf Acc = "COST_STOCK-AMOUNT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "ปรับยอด STOCK เป็นชุด")
      
   ElseIf Acc = "MASTER_REPORT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "รายงานข้อมูลหลัก")
   ElseIf Acc = "MAIN_REPORT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "รายงานข้อมูลส่วนกลาง")
   ElseIf Acc = "PRODUCT_REPORT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "รายงานการผลิต")
   ElseIf Acc = "INVENTORY_REPORT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "รายงานคลัง")
   ElseIf Acc = "LEDGER_REPORT" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "รายงานการบัญชี")
   
   ElseIf Acc = "PROGRAM" Then
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", "โปรแกรม")
      
      
   Else
      Call Ri.SetFieldValue("RIGHT_ITEM_DESC", ReportName)
   End If
   
End Sub
Private Function CreatePermissionNode(Acc As String, ParentID As Long, ReportName As String) As Boolean
Dim ParentKey As String
Dim TopFlag As Boolean
Dim TempParentID As Long
Dim CreateFlag As Boolean
Dim Ri As CRightItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long
   
   'Create node here
   Set Ri = New CRightItem
   Set TempRs = New ADODB.Recordset
   TempParentID = 0
   
   Call Ri.SetFieldValue("RIGHT_ID", -1)
   Call Ri.SetFieldValue("RIGHT_ITEM_NAME", Acc)
   Call Ri.QueryData(1, TempRs, iCount)
   If TempRs.EOF Then
      ParentKey = GetParentKey(Acc, TopFlag)
      If Not TopFlag Then
         Call CreatePermissionNode(ParentKey, TempParentID, ReportName)
         Call Ri.SetFieldValue("PARENT_ID", TempParentID)
      End If
      
      Ri.ShowMode = SHOW_ADD
      Call GetParentItemDesc(Acc, Ri, ReportName)
      Call Ri.AddEditData
      ParentID = Ri.GetFieldValue("RIGHT_ID")
   Else
      Call Ri.PopulateFromRS(1, TempRs)
      ParentID = Ri.GetFieldValue("RIGHT_ID")
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set Ri = Nothing
End Function
Public Function VerifyAccessRight(Acc As String, Optional ReportName As String = "") As Boolean
Dim R As CGroupRight
Dim iCount As Long
Dim TempParentID As Long
Dim FoundFlag As Boolean
   
   If glbUser.REAL_USER_ID = 0 Then
      VerifyAccessRight = True
      Exit Function
   End If
   
   Call glbDaily.StartTransaction
   Call CreatePermissionNode(Acc, TempParentID, ReportName)
   Call glbDaily.CommitTransaction
   
   FoundFlag = False
   For Each R In glbAccessRight
      If R.GetFieldValue("RIGHT_ITEM_NAME") = Acc Then
         FoundFlag = True
         If R.GetFieldValue("RIGHT_STATUS") = "Y" Then
            VerifyAccessRight = True
            Set R = Nothing
            Exit For
         Else
            VerifyAccessRight = False
            Set R = Nothing
            Exit For
         End If
      End If
   Next R
   
   If (Not FoundFlag) Or (Not VerifyAccessRight) Then
      VerifyAccessRight = False
      glbErrorLog.LocalErrorMsg = "ไม่สามารถใช้งานโปรแกรมส่วนนี้ได้เนื่องจากมีสิทธ์ไม่พอเพียง -> " & Acc
      glbErrorLog.ShowUserError
   Else
      VerifyAccessRight = True
   End If
   Set R = Nothing
End Function

Public Sub ReleaseAll()
   Set glbErrorLog = Nothing
   Set glbDatabaseMngr = Nothing
   Set glbParameterObj = Nothing
   Set glbUser = Nothing
   Set glbGuiConfigs = Nothing
   
   Set glbEnterPrise = Nothing
   Set glbDaily = Nothing
   Set glbSetting = Nothing
   Set glbAccessRight = Nothing
   
   Set LoadPackageColl = Nothing
   Set m_CustomerColl = Nothing
   Set m_SupplierColl = Nothing
   Set m_EmployeeColl = Nothing
   Set m_LocationColl = Nothing
   Set InventorySubTypecoll = Nothing
   Set glbLockDate = Nothing
End Sub
Public Function DateToStringExtEx3(D As Date) As String
   If D > 0 Then
      DateToStringExtEx3 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
      DateToStringExtEx3 = DateToStringExtEx3 & " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   Else
      DateToStringExtEx3 = ""
   End If
End Function

Public Function EmptyToLong(V As Variant) As Long
   If V Is Empty Then
      EmptyToLong = 0
   End If
End Function

Public Function VerifyGrid(S As String) As Boolean
   If S = "" Then
      VerifyGrid = False
      glbErrorLog.LocalErrorMsg = "กรุณาเลือกข้อมูลที่ต้องการก่อน"
      glbErrorLog.ShowUserError
   Else
      VerifyGrid = True
   End If
End Function

Public Function VerifyTextControl(L As Label, T As uctlTextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextControl = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextControl = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextControl = True 'false
               Exit Function 'remove this if false
            Else
               VerifyTextControl = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextControl = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextControl = True 'false
                  Exit Function
               Else
                  VerifyTextControl = True
                  Exit Function
               End If
            End If
'            Exit Function
         End If
      End If
      
      VerifyTextControl = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextControl = False
         Exit Function
      End If
      
      VerifyTextControl = True
   End If
End Function

Public Function CountItem(Col As Collection) As Long
Dim I As Long
Dim Count As Long

   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
   Next I
   
   CountItem = Count
End Function

Public Function ConfirmDelete(S As String) As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการจะลบข้อมูล " & S & " ใช่หรือไม่"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmDelete = False
      Exit Function
   Else
      ConfirmDelete = True
   End If
End Function

Public Function GetItem(Col As Collection, Idx As Long, RealIndex As Long) As Object
Dim I As Long
Dim Count As Long

   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
      If Count = Idx Then
         RealIndex = I
         Set GetItem = Col.Item(I)
         Exit Function
      End If
   Next I
   
   Set GetItem = Nothing
End Function

Public Sub InitCheckBox(C As SSCheck, Caption As String)
   C.Caption = Caption
   C.FontSize = 14
   C.FontBold = True
   C.FontName = GLB_FONT
   C.BackColor = GLB_FORM_COLOR
   C.BackStyle = ssTransparent
   C.TripleState = True
End Sub
Public Function FlagToCheck(F As String) As Long
   If F = "Y" Then
      FlagToCheck = 1
   Else
      FlagToCheck = 0
   End If
End Function
Public Function Check2Flag(A As Long) As String
   If A = ssCBChecked Then
      Check2Flag = "Y"
   Else
      Check2Flag = "N"
   End If
End Function

Public Function VerifyCombo(L As Label, C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyCombo = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled And C.Visible Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyCombo = True
End Function
Public Function DateToStringExtEx2(D As Date) As String
   If D > 0 Then
      DateToStringExtEx2 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      DateToStringExtEx2 = ""
   End If
End Function
Public Function DateToStringExtEx(D As Date) As String
   If D < 0 Then
      DateToStringExtEx = ""
      Exit Function
   End If
   
   DateToStringExtEx = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
End Function

Public Function VerifyDate(L As Label, D As uctlDate, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not D.VerifyDate(NullAllow) Then
      VerifyDate = False
      D.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyDate = True
   End If
End Function
Public Function VerifyTime(L As Label, T As uctlTime, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not T.VerifyTime(NullAllow) Then
      VerifyTime = False
      T.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyTime = True
   End If
End Function


Public Function DateToStringIntHi(D As Date) As String
   If D > 0 Then
      DateToStringIntHi = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
   Else
      DateToStringIntHi = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   End If
End Function

Public Function DateToStringIntLow(D As Date) As String
   If D = -1 Then
      DateToStringIntLow = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   ElseIf D = -2 Then
      DateToStringIntLow = "0000" & "-" & "00" & "-" & "00" & _
                     " 00:00:00"
   Else
      DateToStringIntLow = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                        " 00:00:00"
   End If
End Function

Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function

Public Function Doctype2Text(ID As INVENTORY_DOCTYPE) As String
   If ID = IMPORT_DOCTYPE Then
      Doctype2Text = "เอกสารการนำเข้าสต็อค"
   ElseIf ID = EXPORT_DOCTYPE Then
      Doctype2Text = "เอกสารการเบิกจ่ายสต็อค"
   ElseIf ID = TRANSFER_DOCTYPE Then
      Doctype2Text = "เอกสารการโอนสต็อค"
   ElseIf ID = ADJUST_DOCTYPE Then
      Doctype2Text = "เอกสารการปรับยอดสต็อค"
   ElseIf ID = 1000 Then
      Doctype2Text = "เอกสารการผลิต"
   End If
End Function

Public Function ChequeType2Text(ID As Long) As String
   If ID = 1 Then
      ChequeType2Text = "เช็ครับ"
   ElseIf ID = 2 Then
      ChequeType2Text = "เช็คจ่าย"
   End If
End Function

Public Function SellDoctype2Text(ID As SELL_BILLING_DOCTYPE) As String
   If ID = BILLS_DOCTYPE Or ID = S_BILLS_DOCTYPE Then
      SellDoctype2Text = "ใบสรุปวางบิล"
   ElseIf ID = CN_DOCTYPE Or ID = S_CN_DOCTYPE Then
      SellDoctype2Text = "ใบลดหนี้"
   ElseIf ID = DN_DOCTYPE Or ID = S_DN_DOCTYPE Then
      SellDoctype2Text = "ใบเพิ่มหนี้"
   ElseIf ID = INVOICE_DOCTYPE Then
      SellDoctype2Text = "ใบส่งของ/ใบกำกับภาษี"
   ElseIf ID = S_INVOICE_DOCTYPE Then
      SellDoctype2Text = "ใบรับสินค้า"
   ElseIf ID = PO_DOCTYPE Then
      SellDoctype2Text = "ใบสั่งขาย"
   ElseIf ID = S_PO_DOCTYPE Then
      SellDoctype2Text = "ใบสั่งซื้อ"
   
   ElseIf ID = QUOATATION_DOCTYPE Or ID = S_QUOATATION_DOCTYPE Then
      SellDoctype2Text = "ใบเสนอราคา"
   ElseIf ID = RECEIPT1_DOCTYPE Or ID = S_RECEIPT1_DOCTYPE Then
      SellDoctype2Text = "ใบเสร็จรับเงิน (ขายสด)"
   ElseIf ID = RECEIPT2_DOCTYPE Or ID = S_RECEIPT2_DOCTYPE Then
      SellDoctype2Text = "ใบเสร็จรับเงิน (รับชำระ)"
   ElseIf ID = RETURN_DOCTYPE Then
      SellDoctype2Text = "ใบรับคืนสินค้า"
   ElseIf ID = S_RETURN_DOCTYPE Then
      SellDoctype2Text = "ใบส่งคืนสินค้า"
   ElseIf ID = RECEIPT3_DOCTYPE Then
      SellDoctype2Text = "ใบเสร็จรับเงิน (รับชำระเป็นชุด)"
   End If
End Function
Public Function SellDoctype2Report(ID As SELL_BILLING_DOCTYPE) As String
   If ID = BILLS_DOCTYPE Or ID = S_BILLS_DOCTYPE Then
      SellDoctype2Report = "ใบวางบิล"
   ElseIf ID = CN_DOCTYPE Or ID = S_CN_DOCTYPE Then
      SellDoctype2Report = "ใบลดหนี้"
   ElseIf ID = DN_DOCTYPE Or ID = S_DN_DOCTYPE Then
      SellDoctype2Report = "ใบเพิ่มหนี้"
   ElseIf ID = INVOICE_DOCTYPE Then
      SellDoctype2Report = "ใบส่งของ/ใบกำกับภาษี"
   ElseIf ID = S_INVOICE_DOCTYPE Then
      SellDoctype2Report = "ใบรับสินค้า"
   ElseIf ID = PO_DOCTYPE Or ID = S_CN_DOCTYPE Then
      SellDoctype2Report = "ใบสั่งซื้อ"
   ElseIf ID = QUOATATION_DOCTYPE Or ID = S_QUOATATION_DOCTYPE Then
      SellDoctype2Report = "ใบเสนอราคา"
   ElseIf ID = RECEIPT1_DOCTYPE Then
      SellDoctype2Report = "ใบเสร็จรับเงิน"
   ElseIf ID = RECEIPT2_DOCTYPE Then
      SellDoctype2Report = "ใบเสร็จรับเงิน"
   
   ElseIf ID = S_RECEIPT1_DOCTYPE Then
      SellDoctype2Report = "ใบสำคัญจ่าย"
   ElseIf ID = S_RECEIPT2_DOCTYPE Then
      SellDoctype2Report = "ใบสำคัญจ่าย"
   
   
   ElseIf ID = RETURN_DOCTYPE Then
      SellDoctype2Report = "ใบลดหนี้รับคืนสินค้า"
   ElseIf ID = S_RETURN_DOCTYPE Then
      SellDoctype2Report = "ใบลดหนี้ส่งคืนสินค้า"
   End If
End Function
Public Function SellDoctype2ReportEx(ID As SELL_BILLING_DOCTYPE) As String
   If ID = BILLS_DOCTYPE Then
      SellDoctype2ReportEx = "วางบิล"
   ElseIf ID = CN_DOCTYPE Then
      SellDoctype2ReportEx = "ลดหนี้"
   ElseIf ID = DN_DOCTYPE Then
      SellDoctype2ReportEx = "เพิ่มหนี้"
   ElseIf ID = INVOICE_DOCTYPE Then
      SellDoctype2ReportEx = "ขายเชื่อได้"
   ElseIf ID = PO_DOCTYPE Then
      SellDoctype2ReportEx = "สั่งซื้อ"
   ElseIf ID = QUOATATION_DOCTYPE Then
      SellDoctype2ReportEx = "เสนอราคา"
   ElseIf ID = RECEIPT1_DOCTYPE Then
      SellDoctype2ReportEx = "เงินสด"
   ElseIf ID = RECEIPT2_DOCTYPE Then
      SellDoctype2ReportEx = "เงินเชื่อ"
   ElseIf ID = RETURN_DOCTYPE Then
      SellDoctype2ReportEx = "ลดหนี้/รับคืน"
   End If
End Function

Public Function MyDiffEx(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiffEx = 0
   Else
      MyDiffEx = D1 / D2
   End If
End Function
Public Function GetObject(ClassName As String, m_TempCol As Collection, TempKey As String, Optional SetNew As Boolean = True) As Object
On Error Resume Next
Dim Ei As Object
   
   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If SetNew Then
         Set GetObject = GetNewClass(ClassName)
      End If
   Else
      Set GetObject = Ei
   End If
End Function
Public Function GetNewClass(ClassName As String) As Object
   If ClassName = "CStockCode" Then
      Static m_CStockCode As CStockCode
      If m_CStockCode Is Nothing Then
         Set m_CStockCode = New CStockCode
      End If
      Set GetNewClass = m_CStockCode
   
   ElseIf ClassName = "CBillingDoc" Then
      Static m_CBillingDoc As CBillingDoc
      If m_CBillingDoc Is Nothing Then
         Set m_CBillingDoc = New CBillingDoc
      End If
      Set GetNewClass = m_CBillingDoc
      
   ElseIf ClassName = "CDocItem" Then
      Static m_CDocItem As CDocItem
      If m_CDocItem Is Nothing Then
         Set m_CDocItem = New CDocItem
      End If
      Set GetNewClass = m_CDocItem
         
   ElseIf ClassName = "CExportID" Then
      Static m_CExportID As CExportID
      If m_CExportID Is Nothing Then
         Set m_CExportID = New CExportID
      End If
      Set GetNewClass = m_CExportID
   
   ElseIf ClassName = "CCreditBalanceID" Then
      Static m_CCreditBalanceID As CCreditBalanceID
      If m_CCreditBalanceID Is Nothing Then
         Set m_CCreditBalanceID = New CCreditBalanceID
      End If
      Set GetNewClass = m_CCreditBalanceID
      
   ElseIf ClassName = "CAddress" Then
      Static m_CAddress As CAddress
      If m_CAddress Is Nothing Then
         Set m_CAddress = New CAddress
      End If
      Set GetNewClass = m_CAddress
      
   ElseIf ClassName = "CRcpCnDn_Item" Then
      Static m_CRcpCnDn_Item As CRcpCnDn_Item
      If m_CRcpCnDn_Item Is Nothing Then
         Set m_CRcpCnDn_Item = New CRcpCnDn_Item
      End If
      Set GetNewClass = m_CRcpCnDn_Item
   ElseIf ClassName = "CBillingAddition" Then
      Static m_CBillingAddition As CBillingAddition
      If m_CBillingAddition Is Nothing Then
         Set m_CBillingAddition = New CBillingAddition
      End If
      Set GetNewClass = m_CBillingAddition
   ElseIf ClassName = "CBillingSubTract" Then
      Static m_CBillingSubTract As CBillingSubTract
      If m_CBillingSubTract Is Nothing Then
         Set m_CBillingSubTract = New CBillingSubTract
      End If
      Set GetNewClass = m_CBillingSubTract
   
   ElseIf ClassName = "CBalanceVerifyDeTail" Then
      Static m_CBalanceVerifyDeTail As CBalanceVerifyDeTail
      If m_CBalanceVerifyDeTail Is Nothing Then
         Set m_CBalanceVerifyDeTail = New CBalanceVerifyDeTail
      End If
      Set GetNewClass = m_CBalanceVerifyDeTail
   ElseIf ClassName = "CLotItem" Then
      Static m_CLotItem As CLotItem
      If m_CLotItem Is Nothing Then
         Set m_CLotItem = New CLotItem
      End If
      Set GetNewClass = m_CLotItem
   ElseIf ClassName = "CCashTran" Then
      Static m_CCashTran As CCashTran
      If m_CCashTran Is Nothing Then
         Set m_CCashTran = New CCashTran
      End If
      Set GetNewClass = m_CCashTran
   
   ElseIf ClassName = "CMasterRef" Then
      Static m_CMasterRef As CMasterRef
      If m_CMasterRef Is Nothing Then
         Set m_CMasterRef = New CMasterRef
      End If
      Set GetNewClass = m_CMasterRef
   ElseIf ClassName = "CTotalSale" Then
      Static m_CTotalSale As CTotalSale
      If m_CTotalSale Is Nothing Then
         Set m_CTotalSale = New CTotalSale
      End If
      Set GetNewClass = m_CTotalSale
   
   
   End If
   
End Function
Public Function AdjustType2Code(TempID As Long) As String
   If TempID = 1 Then
      AdjustType2Code = "E"
   ElseIf TempID = 2 Then
      AdjustType2Code = "I"
   End If
End Function

Public Function Code2AdjustType(Cd As String) As Long
   If Cd = "E" Then
      Code2AdjustType = 1
   ElseIf Cd = "I" Then
      Code2AdjustType = 2
   End If
End Function
Public Sub InitFormHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 20
   L.FontName = GLB_FONT
   L.Alignment = 2
   L.ForeColor = RGB(0, 10, 0)
End Sub
Public Sub InitOrientation(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2Orientation(orLandscape))
   C.ItemData(1) = orLandscape

   C.AddItem (ID2Orientation(orPortrait))
   C.ItemData(2) = orPortrait
End Sub
Public Function ID2Orientation(TempID As OrientationSettings) As String
   If TempID = orLandscape Then
      ID2Orientation = "แนวนอน"
   Else
      ID2Orientation = "แนวตั้ง"
   End If
End Function
Public Function ID2PaperSize(TempID As PaperSizeSettings) As String
   If TempID = pprA4 Then
      ID2PaperSize = "A4"
   ElseIf TempID = pprLetter Then
      ID2PaperSize = "Letter"
   ElseIf TempID = pprFanfoldUS Then
      ID2PaperSize = "Us standard"
   ElseIf TempID = pprA5 Then
      ID2PaperSize = "A5"
   Else
      ID2PaperSize = "A4"
   End If
End Function
Public Sub InitPaperSize(C As ComboBox)
Dim Vsp As VSPrinter
   
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2PaperSize(pprA4))
   C.ItemData(1) = pprA4

   C.AddItem (ID2PaperSize(pprLetter))
   C.ItemData(2) = pprLetter

   C.AddItem (ID2PaperSize(pprFanfoldUS))
   C.ItemData(3) = pprFanfoldUS
   
    C.AddItem (ID2PaperSize(pprA5))
   C.ItemData(4) = pprA5
   
  
End Sub
Public Sub InitFontName(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("AngsanaUPC")
   C.ItemData(1) = 1
End Sub
Public Function VSP_CalTable(ByVal pRaw As String, ByVal pWidth As Long, ByRef pPer() As Long) As String
On Error GoTo ErrorHandler
Dim strTemp As String
Dim I As Long
Dim Count As Long
Dim iPer As Long
Dim tPer As Long
Dim Total As Long
Dim Prefix() As String
Dim Value() As Long
Dim iTemp As Long
   
   pRaw = Trim$(pRaw)
   If Len(pRaw) <= 0 Then
      VSP_CalTable = ""
      Exit Function
   End If
   Count = 0
   iPer = 1
   Total = 0
   strTemp = ""
   While iPer <= Len(pRaw)
      If Val(Mid$(pRaw, iPer, 1)) <= 0 Then
         strTemp = strTemp & Mid$(pRaw, iPer, 1)
      Else
         Count = Count + 1
         ReDim Preserve Prefix(Count)
         ReDim Preserve Value(Count)
         Prefix(Count) = strTemp
         tPer = InStr(iPer, pRaw, "|")
         If tPer <= 0 Then tPer = InStr(iPer, pRaw, ";")

         Value(Count) = Val(Mid$(pRaw, iPer, tPer - iPer))
         Total = Total + Value(Count)
         iPer = tPer
         strTemp = ""
      End If
      iPer = iPer + 1
   Wend
   strTemp = ""
   ReDim pPer(Count)
   For I = 1 To Count - 1
      iTemp = CLng((Value(I) * pWidth) / Total)
      strTemp = strTemp & Trim$(Prefix(I)) & Trim$(Str$(iTemp)) & "|"
      If I = 1 Then
         pPer(I - 1) = iTemp
      Else
         pPer(I - 1) = pPer(I - 2) + iTemp
      End If
   Next I
   strTemp = strTemp & Trim$(Prefix(I)) & CLng(((Value(I) * pWidth) / Total)) & ";"
   If I > 1 Then
      iTemp = CLng((Value(I) * pWidth) / Total)
      pPer(I - 1) = pPer(I - 2) + iTemp
   End If
   VSP_CalTable = strTemp

   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function
Public Function SetReportConfig(Vsp As VSPrinter, ReportClassName As String, Optional ReportConfig As CReportConfig = Nothing, Optional Flag As Boolean = True) As Boolean
Dim I As Long
Dim Count As Long
Dim Rp As CReportConfig
Dim TempRs As ADODB.Recordset
Dim Rps As Collection
Dim iCount As Long
   
   If Rps Is Nothing Then
      Set TempRs = New ADODB.Recordset
      
      Set Rps = New Collection
      Set Rp = New CReportConfig
      
      Call Rp.SetFieldValue("REPORT_CONFIG_ID", -1)
      Call Rp.QueryData(1, TempRs, iCount)
      Set Rp = Nothing
      
      While Not TempRs.EOF
         Set Rp = New CReportConfig
         
         Call Rp.PopulateFromRS(1, TempRs)
         Call Rps.add(Rp)
         
         Set Rp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Rp = Nothing
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   SetReportConfig = False
   For Each Rp In Rps
      If Rp.GetFieldValue("REPORT_KEY") = ReportClassName Then
         If Flag Then
            If Rp.GetFieldValue("PAPER_SIZE") > 0 Then
               Vsp.PaperSize = Rp.GetFieldValue("PAPER_SIZE")
            End If
            Vsp.Orientation = Rp.GetFieldValue("Orientation")
            Vsp.MarginBottom = Rp.GetFieldValue("MARGIN_BOTTOM") * 567
            Vsp.MarginLeft = Rp.GetFieldValue("MARGIN_LEFT") * 567
            Vsp.MarginRight = Rp.GetFieldValue("MARGIN_RIGHT") * 567
            Vsp.MarginTop = Rp.GetFieldValue("MARGIN_TOP") * 567
            
            If Rp.GetFieldValue("FONT_SIZE") > 0 Then
               Vsp.FontSize = Rp.GetFieldValue("FONT_SIZE")
            End If
            If Len(Rp.GetFieldValue("FONT_NAME")) > 0 Then
               Vsp.FontName = Rp.GetFieldValue("FONT_NAME")
            End If
         End If
               
         If Not ReportConfig Is Nothing Then
            Set ReportConfig = Rp
         End If
               
         SetReportConfig = True
         Exit Function
      End If
   Next Rp
   Set Rps = Nothing
End Function

Public Sub PatchDB()
Dim p As CPatch
   
   Set p = New CPatch
   
'   If Not p.IsPatch("2014_07_23_5_jill") Then '18
'      Call p.Patch_2014_07_23_5_jill
'   End If
   
   
   
   Set p = Nothing
   
End Sub
Public Function GetRcpCnDn_Item(m_TempCol As Collection, TempKey As String) As CRcpCnDn_Item
On Error Resume Next
Dim Ei As CRcpCnDn_Item
Static TempEi As CRcpCnDn_Item

   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CRcpCnDn_Item
      End If
      Set GetRcpCnDn_Item = TempEi
   Else
      Set GetRcpCnDn_Item = Ei
   End If
End Function
Public Function UpDateBDRcpCnDnItem(m_BillingDoc As CBillingDoc)
Dim Rc As CRcpCnDn_Item
   If m_BillingDoc.RcpCnDnItems.Count > 0 Then
      Dim RcQ As CRcpCnDn_Item
      Dim BiUpDate As CBillingDoc
      Dim RcQColl As Collection
   
      Set RcQ = New CRcpCnDn_Item
      Set RcQColl = New Collection
   
      Call LoadUpdateRcpCnDn(RcQ, Nothing, RcQColl)
   
      Set RcQ = Nothing
   
      For Each Rc In m_BillingDoc.RcpCnDnItems
         Set RcQ = GetRcpCnDn_Item(RcQColl, Trim(Str(Rc.GetFieldValue("DOC_ID"))))
         Set BiUpDate = New CBillingDoc
         
         BiUpDate.BILLING_DOC_ID = Rc.GetFieldValue("DOC_ID")
         BiUpDate.PAID_AMOUNT = RcQ.GetFieldValue("PAID_AMOUNT")
         BiUpDate.PAY_AMOUNT = RcQ.GetFieldValue("ITEM_AMOUNT")
         BiUpDate.UpDatePaidAmount
      Next Rc
   End If
End Function
Public Function GetExportItem(Ivd As CInventoryDoc, GuiID As Long, Optional TxType As String = "") As CLotItem
Dim Ei As CLotItem

   For Each Ei In Ivd.ImportExportItems
      If Ei.LINK_ID = GuiID Then
         If TxType = "" Then
            Set GetExportItem = Ei
            Exit Function
         ElseIf TxType = Ei.TX_TYPE Then
            Set GetExportItem = Ei
            Exit Function
         End If
      End If
   Next Ei
End Function
Public Function MyDiff(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiff = 0
   Else
      MyDiff = CDbl(Format(D1 / D2, "0.00"))
   End If
End Function
Public Function EmptyToString(Value As String, S As String) As String
On Error Resume Next

   If Value = "" Or Value = "0" Or Value = "0.00" Or Value = "0.00%" Or Value = "0%" Then
      EmptyToString = S
   Else
      EmptyToString = Value
   End If
End Function
Public Sub LoadPictureFromFile(FileName As String, Pc As PictureBox)
On Error Resume Next
    If Dir(FileName) <> "" Then
      Pc.Picture = LoadPicture(FileName)
   End If
End Sub
Public Sub GetFirstLastDate(D As Date, FD As Date, LD As Date, Optional add As Long = 0)
Dim MM As Long
Dim DD1 As Long
Dim DD2 As Long
Dim YYYY As Long
   D = DateAdd("m", add, D)
   MM = Month(D)
   DD1 = 1
   DD2 = LastDayOfMonth(D)
   YYYY = Year(D)
   
   FD = DateSerial(YYYY, MM, DD1)
   LD = DateSerial(YYYY, MM, DD2)
End Sub
Public Function GetFirstLastDateEX(F As Date, L As Date, FD As Date, LD As Date) As Long
Dim MM1 As Long
Dim DD1 As Long
Dim YYYY1 As Long

Dim MM2 As Long
Dim DD2 As Long
Dim YYYY2 As Long

   MM1 = Month(F)
   DD1 = 1
   YYYY1 = Year(F)
   
   MM2 = Month(L)
   DD2 = LastDayOfMonth(L)
   YYYY2 = Year(L)
   
   
   FD = DateSerial(YYYY1, MM1, DD1)
   LD = DateSerial(YYYY2, MM2, DD2)
   
   GetFirstLastDateEX = DateDiff("D", FD, LD) + 1
End Function

Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
Dim LastDay As Byte
   LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", ValidDate) + 1, ValidDate))))
   LastDayOfMonth = LastDay
End Function
Public Function AdjustPage(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional TestFlag As Boolean = False, Optional SpaceCount As Long) As Boolean
Dim TempStr As String
   
   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   If (Vsp.Y1 + Offset - SpaceCount) > (Vsp.PageHeight - Vsp.MarginBottom) Then
      If Not TestFlag Then
         Vsp.NewPage
      End If
      AdjustPage = True
   Else
      AdjustPage = False
   End If
End Function

Public Function PatchTable(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional EnableFlag As Boolean = True, Optional SpaceCount As Long = 0) As Boolean
Dim TempStr As String
Dim TempBorder As String
   
   TempBorder = Vsp.TableBorder
   If Not EnableFlag Then
      PatchTable = True
      Exit Function
   End If
   
   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   'Vsp.TableBorder = tbColumns
   While Not AdjustPage(Vsp, Header, Body, Offset, True, SpaceCount)
      Call Vsp.AddTable(Header, "", Body)
   Wend
   Vsp.TableBorder = TempBorder
End Function
Public Function CheckUniqueNs(UnqType As UNIQUE_TYPE, Key As String, ID As Long, Optional FieldNameExTendValue As String, Optional FieldNameExTendValueEX As String, Optional NullFlag As Boolean = False) As Boolean
On Error GoTo ErrorHandler
Dim TableName As String
Dim FieldName1 As String
Dim FieldName2 As String
Dim FieldNameExTend As String
Dim FieldNameExTendEX As String
Dim Flag As Boolean
Dim Count As Long

   CheckUniqueNs = False
   
   Flag = False
   
   If UnqType = PACKAGE_NO Then
      TableName = "PACKAGE"
      FieldName1 = "PACKAGE_NO"
      FieldName2 = "PACKAGE_ID"
      Flag = True
    ElseIf UnqType = PACKAGE_DESC Then
      TableName = "PACKAGE"
      FieldName1 = "PACKAGE_DESC"
      FieldName2 = "PACKAGE_ID"
      Flag = True
   ElseIf UnqType = PACKAGE_MASTER_FLAG Then
      TableName = "PACKAGE"
      FieldName1 = "PACKAGE_MASTER_FLAG"
      FieldName2 = "PACKAGE_ID"
      'FieldNameExTend = "PACKAGE_TYPE"
      Flag = True
   ElseIf UnqType = DOCUMENT_NO_UNIQUE Then
      TableName = "BILLING_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "BILLING_DOC_ID"
      Flag = True
   ElseIf UnqType = APARCODE_UNIQUE Then
      TableName = "APAR_MAS"
      FieldName1 = "APAR_CODE"
      FieldName2 = "APAR_MAS_ID"
      Flag = True
   ElseIf UnqType = INVENTORY_DOC_NO Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = PARTNO_UNIQUE Then
      TableName = "STOCK_CODE"
      FieldName1 = "STOCK_NO"
      FieldName2 = "STOCK_CODE_ID"
      Flag = True
   ElseIf UnqType = MASTER_CODE Then
      TableName = "MASTER_REF"
      FieldName1 = "KEY_CODE"
      FieldName2 = "KEY_ID"
      FieldNameExTend = "MASTER_AREA"
      Flag = True
   ElseIf UnqType = MASTER_NAME Then
      TableName = "MASTER_REF"
      FieldName1 = "KEY_NAME"
      FieldName2 = "KEY_ID"
      FieldNameExTend = "MASTER_AREA"
      Flag = True
   ElseIf UnqType = EMPCODE_UNIQUE Then
      TableName = "EMPLOYEE"
      FieldName1 = "EMP_CODE"
      FieldName2 = "EMP_ID"
      Flag = True
   ElseIf UnqType = MASTER_VALID_NO Then
      TableName = "MASTER_VALID"
      FieldName1 = "MASTER_VALID_NO"
      FieldName2 = "MASTER_VALID_ID"
      FieldNameExTend = "MASTER_VALID_TYPE"
      Flag = True
   ElseIf UnqType = FORMULA_NO Then
      TableName = "FORMULA"
      FieldName1 = "FORMULA_NO"
      FieldName2 = "FORMULA_ID"
      Flag = True
   ElseIf UnqType = JOB_NO Then
      TableName = "JOB"
      FieldName1 = "JOB_NO"
      FieldName2 = "JOB_ID"
      Flag = True
   End If
   
   If Flag Then
      Count = glbDatabaseMngr.CountRecord(TableName, FieldName1, FieldName2, Key, ID, glbErrorLog, FieldNameExTend, FieldNameExTendValue, FieldNameExTendEX, FieldNameExTendValueEX, NullFlag)
      If Count <> 0 Then
         CheckUniqueNs = False
      Else
         CheckUniqueNs = True
      End If
   End If
      
   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   CheckUniqueNs = False
End Function

Public Sub StartExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
   Vsp.ExportFile = glbParameterObj.ReportFile
   Vsp.ExportFormat = vpxPlainHTML
End Sub

Public Sub CloseExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
End Sub
Public Function PatchWildCard(T As String) As String
   If Len(Trim(T)) <> 0 Then
      PatchWildCard = T & "%"
   Else
      PatchWildCard = T
   End If
End Function
Public Sub InitOptionEx(O As SSOption, Caption As String)
   O.Caption = Caption
   O.Font.Size = 14
   O.Font.Bold = True
   O.Font.Name = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
   O.BackStyle = ssTransparent
End Sub
Public Function ConvertDocToConfigNo(DocKind As Long, DocType As SELL_BILLING_DOCTYPE, DocSubType As Long) As Long
   If DocKind = 1 Then
      If DocType = QUOATATION_DOCTYPE Then
         ConvertDocToConfigNo = 1
      ElseIf DocType = PO_DOCTYPE Then
         ConvertDocToConfigNo = 2
      ElseIf DocType = INVOICE_DOCTYPE Then
         ConvertDocToConfigNo = 100 + DocSubType
      ElseIf DocType = RECEIPT1_DOCTYPE Then
         ConvertDocToConfigNo = 3
      ElseIf DocType = RECEIPT2_DOCTYPE Then
         ConvertDocToConfigNo = 4
      ElseIf DocType = DN_DOCTYPE Then
         ConvertDocToConfigNo = 5
      ElseIf DocType = CN_DOCTYPE Then
         ConvertDocToConfigNo = 6
      ElseIf DocType = RETURN_DOCTYPE Then
         ConvertDocToConfigNo = 7
      ElseIf DocType = BILLS_DOCTYPE Then
         ConvertDocToConfigNo = 8
      ElseIf DocType = S_QUOATATION_DOCTYPE Then
         ConvertDocToConfigNo = 19
      ElseIf DocType = S_PO_DOCTYPE Then
         ConvertDocToConfigNo = 20
      ElseIf DocType = S_INVOICE_DOCTYPE Then
         ConvertDocToConfigNo = 21
      ElseIf DocType = S_RECEIPT1_DOCTYPE Then
         ConvertDocToConfigNo = 22
      ElseIf DocType = S_RECEIPT2_DOCTYPE Then
         ConvertDocToConfigNo = 23
      ElseIf DocType = S_DN_DOCTYPE Then
         ConvertDocToConfigNo = 24
      ElseIf DocType = S_CN_DOCTYPE Then
         ConvertDocToConfigNo = 25
      ElseIf DocType = S_RETURN_DOCTYPE Then
         ConvertDocToConfigNo = 26
      ElseIf DocType = S_BILLS_DOCTYPE Then
         ConvertDocToConfigNo = 27
      End If
   ElseIf DocKind = 2 Then
      If DocType = IMPORT_DOCTYPE Then
         ConvertDocToConfigNo = 50
      ElseIf DocType = EXPORT_DOCTYPE Then
         ConvertDocToConfigNo = 51
      ElseIf DocType = TRANSFER_DOCTYPE Then
         ConvertDocToConfigNo = 52
      ElseIf DocType = ADJUST_DOCTYPE Then
         ConvertDocToConfigNo = 53
      ElseIf DocType = 1000 Then
         ConvertDocToConfigNo = 1000
      End If
   ElseIf DocKind = 3 Then
      ConvertDocToConfigNo = 9 'ใบเสร็จ แนบใบส่งของ
   End If
End Function
Public Sub GetCreditBalance(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim BD As CBillingDoc
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CBillingDoc
Dim I As Long
   
   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   BD.FROM_DATE = FromDate
   BD.TO_DATE = ToDate
   BD.ORDER_TYPE = 9999
   BD.DOCUMENT_TYPE_SET = "(" & INVOICE_DOCTYPE & "," & RECEIPT2_DOCTYPE & "," & CN_DOCTYPE & "," & DN_DOCTYPE & "," & RETURN_DOCTYPE & ")"
   BD.FROM_APAR_CODE = FromAparCode
   BD.TO_APAR_CODE = ToAparCode
   Call BD.QueryData(12, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CBillingDoc
      Call TempData.PopulateFromRS(12, Rs)
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(TempData.APAR_MAS_ID & "-" & TempData.DOCUMENT_TYPE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function SelectFlagToText(SelectFlag As String) As String
   If SelectFlag = "Y" Then
      SelectFlagToText = "ใช้"
   Else
      SelectFlagToText = "ไม่"
   End If
End Function
Public Function AddStringFrontEnd(Value As String, Optional F As String, Optional E As String) As String
On Error Resume Next
   If Len(Trim(Value)) > 0 Then
      If Len(F) = 0 And Len(E) = 0 Then
         AddStringFrontEnd = "(" & Value & ")"
      Else
         AddStringFrontEnd = " " & F & " " & Value & " " & E & " "
      End If
   End If
End Function
Public Function AddDoubleFrontEnd(Value As Double, Optional F As String, Optional E As String) As String
On Error Resume Next
   If Val(Value) > 0 Then
      If Len(F) = 0 And Len(E) = 0 Then
         AddDoubleFrontEnd = "(" & Value & ")"
      Else
         AddDoubleFrontEnd = " " & F & " " & Value & " " & E & " "
      End If
   End If
End Function
Public Function PaymentTypeToText(ID As PAYMENT_TYPE) As String
   If ID = CASH_PMT Then
      PaymentTypeToText = MapText("เงินสด")
   ElseIf ID = CHEQUE_HAND_PMT Then
      PaymentTypeToText = MapText("เช็คเข้ามือ")
   ElseIf ID = CHEQUE_BANK_PMT Then
      PaymentTypeToText = MapText("เช็คเข้าธนาคาร")
   ElseIf ID = BANKTRF_PMT Then
      PaymentTypeToText = MapText("โอนผ่านธนาคาร")
   End If
End Function

Public Sub InitPaymentType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (PaymentTypeToText(CASH_PMT))
   C.ItemData(1) = CASH_PMT
   
   C.AddItem (PaymentTypeToText(CHEQUE_HAND_PMT))
   C.ItemData(2) = CHEQUE_HAND_PMT
   
   C.AddItem (PaymentTypeToText(CHEQUE_BANK_PMT))
   C.ItemData(3) = CHEQUE_BANK_PMT
   
   C.AddItem (PaymentTypeToText(BANKTRF_PMT))
   C.ItemData(4) = BANKTRF_PMT
End Sub
Public Sub GetPaidAmountByDocID(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional DocumentType As Long)
On Error GoTo ErrorHandler       'ยอดยกมา ของ ลูกหนี้
Dim Rcp As CRcpCnDn_Item
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRcpCnDn_Item
Dim I As Long
Dim TempRcp As Double
Dim TempCreDit As Double
   
   TempRcp = 0
   TempCreDit = 0
   
   Set Rcp = New CRcpCnDn_Item
   Set Rs = New ADODB.Recordset
   
   Rcp.FROM_DATE = FromDate
   Rcp.TO_DATE = ToDate
   Rcp.FROM_APAR_CODE = FromAparCode
   Rcp.TO_APAR_CODE = ToAparCode
   Rcp.DOCUMENT_TYPE = DocumentType
   Call Rcp.QueryDataReport(3, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CRcpCnDn_Item
      Call TempData.PopulateFromRS(3, Rs)
      
      If TempData.DOCUMENT_TYPE = INVOICE_DOCTYPE Then
         TempRcp = TempRcp + TempData.PAID_AMOUNT
      ElseIf TempData.DOCUMENT_TYPE = RETURN_DOCTYPE Then
         TempCreDit = TempCreDit + TempData.PAID_AMOUNT
      Else
         'debug.print
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.add(TempData, Trim(Str(TempData.DOC_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
'   'debug.print (TempRcp)
'   'debug.print (TempCreDit)
   
   
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub GetPaidAmountBySaleCode(Cl As Collection, Optional FromDate As Date, Optional ToDate As Date, Optional FromAparCode As String, Optional ToAparCode As String, Optional FromSaleCode As String, Optional ToSaleCode As String, Optional DocumentType As Long)
On Error GoTo ErrorHandler
Dim Rcp As CRcpCnDn_Item
Dim ItemCount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CRcpCnDn_Item
Dim I As Long

   
   Set Rcp = New CRcpCnDn_Item
   Set Rs = New ADODB.Recordset

'   Rcp.FROM_DATE = FromDate
   Rcp.TO_DATE = ToDate
   Rcp.FROM_APAR_CODE = FromAparCode
   Rcp.TO_APAR_CODE = ToAparCode
   Rcp.DOCUMENT_TYPE = DocumentType
   Call Rcp.QueryDataReport(5, Rs, ItemCount)
   
   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      Set TempData = New CRcpCnDn_Item
      Call TempData.PopulateFromRS(5, Rs)
      
      If Not (Cl Is Nothing) Then
'         Call Cl.add(TempData, Trim(Str(TempData.DOC_ID) & "-" & Str(TempData.DOCUMENT_TYPE) & "-" & TempData.APAR_CODE))
         Call Cl.add(TempData, Trim(Str(TempData.DOC_ID) & "-" & TempData.APAR_CODE))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend

   
   Set Rcp = Nothing
   Set Rs = Nothing
   Exit Sub
   
ErrorHandler:
   
   glbErrorLog.SystemErrorMsg = Err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Function StringToDateCheckError(S As String) As Date
On Error Resume Next
   StringToDateCheckError = -1
   StringToDateCheckError = S
End Function
Public Function CheckSSoptionToString(A As Boolean) As String
   If A = True Then
      CheckSSoptionToString = "Y"
   Else
      CheckSSoptionToString = "N"
   End If
End Function
Public Function StringToCheckSSoption(A As String) As Boolean
   If A = "Y" Then
      StringToCheckSSoption = True
   Else
      StringToCheckSSoption = False
   End If
End Function
Public Function CheckHaveValue(OldCheckHaveValue As Boolean, Amt As Double) As Boolean
   If (Amt <> 0) Or OldCheckHaveValue Then
      CheckHaveValue = True
   End If
End Function
Public Function GetDatePeriodString(FromDate As Date, ToDate As Date) As String
Dim TempStringFrom As String
Dim TempStringTo As String
   TempStringFrom = DateToStringExtEx2(FromDate)
   TempStringTo = DateToStringExtEx2(ToDate)
   If FromDate = ToDate Then
      GetDatePeriodString = TempStringFrom
   ElseIf (Year(FromDate) = Year(ToDate)) And (Month(FromDate) = Month(ToDate)) Then
      GetDatePeriodString = Left(TempStringFrom, 2) & "-" & Left(TempStringTo, 2) & "/" & Mid(TempStringFrom, 4, 2) & "/" & Right(TempStringFrom, 4)
   ElseIf (Year(FromDate) = Year(ToDate)) Then
      GetDatePeriodString = Left(TempStringFrom, 2) & "/" & Mid(TempStringFrom, 4, 2) & "-" & Left(TempStringTo, 2) & "/" & Mid(TempStringTo, 4, 2) & "/" & Right(TempStringFrom, 4)
   Else
      GetDatePeriodString = TempStringFrom & "-" & TempStringTo
   End If
End Function
Public Sub UnLoadAllForm()
Dim F As Form
   For Each F In Forms
      If F.Name <> frmWinPricingMain.Name Then
         Unload F
         Set F = Nothing
      End If
   Next F


End Sub
Public Sub LoadCalculator()
On Error Resume Next
   Call Shell("C:\WINDOWS\system32\calc.exe ", vbMaximizedFocus)
End Sub
Public Function StringToFreeFlag(I As Long) As String
If I > 0 Then
   StringToFreeFlag = ""
Else
   StringToFreeFlag = "N"
End If
End Function
Public Function VerifyLockDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_DATE And uctlDate <= glbLockDate.TO_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_DATE And oldDate <= glbLockDate.TO_DATE))) Then
      VerifyLockDate = True
   Else
      VerifyLockDate = False
   End If
End Function
Public Function VerifyLockInventoryDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_INVENTORY_DATE And uctlDate <= glbLockDate.TO_INVENTORY_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_INVENTORY_DATE And oldDate <= glbLockDate.TO_INVENTORY_DATE))) Then
      VerifyLockInventoryDate = True
   Else
      VerifyLockInventoryDate = False
   End If
End Function
Public Function VerifyLockInvoiceDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_INVOICE_DATE And uctlDate <= glbLockDate.TO_INVOICE_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_INVOICE_DATE And oldDate <= glbLockDate.TO_INVOICE_DATE))) Then
      VerifyLockInvoiceDate = True
   Else
      VerifyLockInvoiceDate = False
   End If
End Function
Public Function VerifyLockReceiptDate(uctlDate As Date, oldDate As Date) As Boolean
   If (uctlDate >= glbLockDate.FROM_RECEIPT_DATE And uctlDate <= glbLockDate.TO_RECEIPT_DATE And (oldDate <= 0 Or (oldDate >= glbLockDate.FROM_RECEIPT_DATE And oldDate <= glbLockDate.TO_RECEIPT_DATE))) Then
      VerifyLockReceiptDate = True
   Else
      VerifyLockReceiptDate = False
   End If
End Function
Public Function DealerTypeToString(ID As DEALER_TYPE_AREA) As String
   If ID <= 0 Then
      DealerTypeToString = ""
   ElseIf ID = SILVER Then
      DealerTypeToString = "SILVER"
   ElseIf ID = SILVER_PLUS Then
      DealerTypeToString = "SILVER+"
   ElseIf ID = SILVER_PLUS_PLUS Then
      DealerTypeToString = "SILVER++"
   ElseIf ID = GOLD_MUNUS Then
      DealerTypeToString = "GOLD-"
   ElseIf ID = GOLD Then
      DealerTypeToString = "GOLD"
   ElseIf ID = GOLD_PLUS Then
      DealerTypeToString = "GOLD+"
   ElseIf ID = GOLD_PLUS_PLUS Then
      DealerTypeToString = "GOLD++"
   ElseIf ID = PLATINUM_MUNUS Then
      DealerTypeToString = "PLATINUM-"
   ElseIf ID = PLATINUM Then
      DealerTypeToString = "PLATINUM"
   ElseIf ID = HEADER_GROUP Then
      DealerTypeToString = "HEADER_GROUP"
   End If
End Function
Public Function DealerTypeToPercent(ID As DEALER_TYPE_AREA) As Double
   If ID <= 0 Then
      DealerTypeToPercent = 0
   ElseIf ID = SILVER Then
      DealerTypeToPercent = 5
   ElseIf ID = SILVER_PLUS Then
      DealerTypeToPercent = 5
   ElseIf ID = SILVER_PLUS_PLUS Then
      DealerTypeToPercent = 5
   ElseIf ID = GOLD_MUNUS Then
      DealerTypeToPercent = 7
   ElseIf ID = GOLD Then
      DealerTypeToPercent = 7
   ElseIf ID = GOLD_PLUS Then
      DealerTypeToPercent = 7
   ElseIf ID = GOLD_PLUS_PLUS Then
      DealerTypeToPercent = 7
   ElseIf ID = PLATINUM_MUNUS Then
      DealerTypeToPercent = 9
   ElseIf ID = PLATINUM Then
      DealerTypeToPercent = 9
   ElseIf ID = HEADER_GROUP Then
      DealerTypeToPercent = 0
   End If
End Function

