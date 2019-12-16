VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCreditStockCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CREDIT_BALANCE As Double

Public TOTAL_PRICE As Double
Public VAT_AMOUNT As Double
Public DISCOUNT_AMOUNT As Double
Public EXT_DISCOUNT_AMOUNT As Double
Public AVG_PRICE As Double

Public TOTAL_AMOUNT As Double
Public STOCK_NO As String
Public STOCK_DESC As String
Public UNIT_CHANGE_NAME As String
Public UNIT_NAME As String
      
Public UNIT_AMOUNT As Double

Public APAR_CODE As String
Public APAR_NAME As String

Public CUSTOMER_BRANCH As Long
Public CUSTOMER_BRANCH_CODE As String
Public CUSTOMER_BRANCH_NAME As String
