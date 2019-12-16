VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExportID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public INVENTORY_DOC_ID As Long
Public LINK_ID As Long

Public AVG_PRICE As Double
Public Function GetKey1() As String
   GetKey1 = Trim(INVENTORY_DOC_ID & "-" & LINK_ID)
End Function

