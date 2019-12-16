VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCreditBalanceID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public APAR_MAS_ID As Long
Public APAR_CODE As String
Public YYYYMM As String

Public CREDIT_BALANCE As Double

Public TOTAL_PRICE As Double
Public VAT_AMOUNT As Double
Public DISCOUNT_AMOUNT As Double
Public EXT_DISCOUNT_AMOUNT As Double

Public PAID_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public DEBIT_AMOUNT As Double
Public SUBTRACT_AMOUNT As Double

Public TOTAL_INVOICE As Double
Public TOTAL_RCP As Double

Public TOTAL_AMOUNT As Double
Public STOCK_NO As String


Public Function GetKey1() As String
   GetKey1 = Trim(Str(APAR_MAS_ID))
End Function
Public Function GetKey2() As String
   GetKey2 = Trim(APAR_MAS_ID & "-" & YYYYMM)
End Function

