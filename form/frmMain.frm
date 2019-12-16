VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3300
      TabIndex        =   0
      Top             =   210
      Width           =   1185
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset

Private Sub Command1_Click()
Dim Sp As CSystemParam
Dim iCount As Long

   Set Sp = New CSystemParam
   Call Sp.RegisterFields

   Call Sp.SetFieldValue("PARAM_ID", 2)
   Call Sp.SetFieldValue("ORDER_BY", 1)
   Call Sp.SetFieldValue("ORDER_TYPE", 1)
   
   Call Sp.QueryData(1, m_Rs, iCount)
   While Not m_Rs.EOF
      Call Sp.PopulateFromRs(1, m_Rs)
'Debug.Print Sp.GetFieldValue("PARAM_ID") & " " & Sp.GetFieldValue("PARAM_NAME") & " " & Sp.GetFieldValue("PARAM_VALUE")
      Call m_Rs.MoveNext
   Wend
   
   Set Sp = Nothing
End Sub

Private Sub Form_Load()
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Rs = Nothing
End Sub
