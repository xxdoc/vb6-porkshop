VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   10515
   ClientLeft      =   1740
   ClientTop       =   555
   ClientWidth     =   13905
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter7LibCtl.VSPrinter VSPrinter1 
      Height          =   9585
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   13935
      _cx             =   24580
      _cy             =   16907
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   52.1390374331551
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   7
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1032
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   585
      Left            =   0
      TabIndex        =   2
      Top             =   9960
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   1032
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   11820
         TabIndex        =   4
         Top             =   40
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   495
         Left            =   9720
         TabIndex        =   3
         Top             =   40
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME = "frmReport"

Private HasActivate As Boolean
Public HeaderText As String
Public ReportID As String
Public ReportObject As CReportInterface
Public OKClick As Boolean
Private m_ErrorFlag As Boolean
Public ClassName As String
Public AutoPrintMode As Boolean

Public Space As Long


Dim g_CharSpacing%

' note: this API is declared incorrectly in the VB API Viewer.
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long


Private Sub cmdPrint_Click()
On Error GoTo ErrorHandler
Dim lMenuChosen As Long
Dim oMenu As CPopupMenu

   Set oMenu = New CPopupMenu
   lMenuChosen = oMenu.Popup("พิมพ์ไปเครื่องพิมพ์", "-", "บันทึกไปที่ไฟล์")
   If lMenuChosen = 0 Then
      Exit Sub
   End If

   If lMenuChosen = 1 Then
      VSPrinter1.PrintDoc (True)
      If m_ErrorFlag Then
         glbErrorLog.LocalErrorMsg = "พบข้อผิดพลาด"
         glbErrorLog.ShowUserError
         Exit Sub
      Else
         glbErrorLog.LocalErrorMsg = "โปรแกรมได้ทำการพิมพ์รายงานเสร็จสิ้นแล้ว"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   ElseIf lMenuChosen = 3 Then
      CommonDialog1.Filter = "Save Files (*.html, *.htm)|*.html;*.htm;"
      CommonDialog1.DialogTitle = "Select access file to import"
      CommonDialog1.ShowSave
      If CommonDialog1.FileName = "" Then
         Exit Sub
      End If
      
      Call FileCopy(glbParameterObj.ReportFile, CommonDialog1.FileName)
   End If
   
   OKClick = True
   Unload Me
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub
Private Sub Form_Activate()
   If Not HasActivate Then
      HasActivate = True
      Me.Refresh
'      DoEvents
      
      Call EnableForm(Me, False)
      
      pnlHeader.Caption = "กรุณารอซักครู่ ระบบกำลังสร้างรายงาน"
      
      Me.Refresh
      Set ReportObject.VsPrint = VSPrinter1
      If Not ReportObject.Preview Then
         glbErrorLog.LocalErrorMsg = ReportObject.ErrorMsg
         glbErrorLog.ShowUserError
      End If
      Call EnableForm(Me, True)
      pnlHeader.Caption = "การสร้างรายงานเสร็จสมบูรณ์"
      
      If AutoPrintMode Then
         Call cmdExit_Click
      End If
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = ClassName
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 121 Then
      Call cmdPrint_Click
      KeyCode = 0
   ElseIf Shift = 1 And KeyCode = 113 Then
      Call Shell("C:\WINDOWS\system32\calc.exe ", vbMaximizedFocus)
   End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

   Me.BackColor = GLB_FORM_COLOR
   VSPrinter1.NavBarColor = GLB_FORM_COLOR
   VSPrinter1.PaperSize = pprA4
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSPanel1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlHeader.Caption = HeaderText
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   HasActivate = False
   m_ErrorFlag = False
   
   Me.Caption = HeaderText
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Form_Load"
   
   Call InitMainButton(cmdPrint, "พิมพ์ (F10)")
   Call InitMainButton(cmdExit, "ออก (ESC)")

   Call EnableForm(Me, True)
   
   g_CharSpacing = Space / VSPrinter1.TwipsPerPixelX
   SetTextCharacterExtra VSPrinter1.hdc, g_CharSpacing
   
   Exit Sub

ErrorHandler:
   Call EnableForm(Me, True)
   glbErrorLog.SystemErrorMsg = err.Description
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set ReportObject = Nothing
   Unload Me
End Sub
Private Sub Form_Resize()
On Error Resume Next
   With VSPrinter1
      .Move .Left, .Top, ScaleWidth - .Left * 2, ScaleHeight - .Top - .Left - 650
      SSPanel1.Top = ScaleHeight - SSPanel1.Height
      SSPanel1.Width = ScaleWidth
      cmdPrint.Left = .Left + ScaleWidth - .Left * 2 - cmdPrint.Width - cmdExit.Width - 20
      cmdExit.Left = .Left + ScaleWidth - .Left * 2 - cmdExit.Width
      pnlHeader.Width = ScaleWidth
      .ZoomMode = zmPageWidth
   End With
End Sub
Private Sub VSPrinter1_NewPage()
    SetTextCharacterExtra VSPrinter1.hdc, g_CharSpacing
End Sub
