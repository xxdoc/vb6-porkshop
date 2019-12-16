VERSION 5.00
Begin VB.Form frmDBSetting 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDBSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Height          =   675
      Left            =   0
      TabIndex        =   6
      Top             =   -210
      Width           =   6225
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   5955
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2580
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   6225
      Begin VB.TextBox txtFileDB 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   420
         Width           =   4455
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   810
         Width           =   4455
      End
      Begin VB.Label lblFileDB 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   300
         TabIndex        =   10
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   330
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   300
         TabIndex        =   8
         Top             =   900
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDBSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKClick As Boolean
Public Header As String

Public FileDb As String
Public UserName As String
Public Password As String
Public IP As String
Public Port As String

Private Sub cmdCancel_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   FileDb = txtFileDB.Text
   UserName = txtUserName.Text
   Password = txtPassword.Text
   
   Call EnableForm(Me, False)
   If Not glbDatabaseMngr.ConnectDatabase(FileDb, UserName, Password, glbErrorLog) Then
      Call EnableForm(Me, True)
      txtFileDB.SetFocus
      Exit Sub
   End If
   
   
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
End Sub

Public Sub InitNormalLabel(L As Label, Caption As String, Optional Color As Long = 0)
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub

Public Sub InitOption(O As OptionButton, Caption As String)
   O.Caption = Caption
   O.FontSize = 14
   O.FontBold = True
   O.FontName = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
End Sub

Public Sub InitOptionEx(O As SSOption, Caption As String)
   O.Caption = Caption
   O.Font.Size = 14
   O.Font.Bold = True
   O.Font.Name = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
   O.BackStyle = ssTransparent
End Sub

Public Sub InitCheckBox(C As SSCheck, Caption As String)
   C.Caption = Caption
   C.FontSize = 14
   C.FontBold = True
   C.FontName = GLB_FONT
   C.BackColor = GLB_FORM_COLOR
   C.BackStyle = ssTransparent
   C.TripleState = True
End Sub

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

Private Sub Form_Load()
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   lblHeader.BackColor = GLB_HEAD_COLOR
   Frame2.BackColor = GLB_HEAD_COLOR

   OKClick = False
'   Call InitDialogHeader(lblHeader, Header)
   
   Call InitNormalLabel(lblFileDB, "Database")
   Call InitNormalLabel(lblUserName, "User name")
   Call InitNormalLabel(lblPassword, "Password")
      
   Call InitTextBox(txtFileDB, FileDb)
   Call InitTextBox(txtUserName, UserName)
   Call InitTextBox(txtPassword, Password, "*")
         
   Call InitDialogButton(cmdOK, "OK")
   Call InitDialogButton(cmdCancel, "CANCEL")
End Sub

