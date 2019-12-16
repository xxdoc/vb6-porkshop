VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTotalCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EMP_ID As Long
Public GROUP_COM_ID As Long
Public TOTAL_PRICE_SUM As Double
Public TOTAL_AMOUNT_SUM As Double

Public TOTAL_PRICE_TOP As Double
Public TOTAL_AMOUNT_TOP As Double

Public TOTAL_AMOUNT_RETURN As Double
Public TOTAL_PRICE_RETURN As Double

Public TAGET_AMOUNT_SUM As Double
Public TAGET_PRICE_SUM As Double
Public GROUP_COM_DESC As String
Public Function Getkey() As String
   Getkey = Trim(EMP_ID & "-" & GROUP_COM_ID)
End Function
