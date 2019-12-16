VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTableField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FieldName As String
Public FieldType As FIELD_TYPE
Public FieldCat As FIELD_CAT

Private FieldValue As Variant
Private m_ADOConn As ADODB.Connection
Public Function GetValue() As Variant
   GetValue = FieldValue
End Function
Public Sub SetValue(Value As Variant)
   FieldValue = Value
End Sub
Public Function TransformToSQLString() As Variant
   If FieldType = STRING_TYPE Then
      TransformToSQLString = "'" & ChangeQuote(Trim(FieldValue)) & "'"
   ElseIf FieldType = DATE_TYPE Then
      TransformToSQLString = "'" & DateToStringInt(Trim(Str(FieldValue))) & "'"
   ElseIf (FieldType = MONEY_TYPE) Then
      TransformToSQLString = FieldValue
   ElseIf (FieldType = INT_TYPE) Then
      If Val(FieldValue) <= 0 Then
         TransformToSQLString = "NULL"
      Else
         TransformToSQLString = Val(FieldValue)
      End If
   Else
      TransformToSQLString = "'" & FieldValue & "'"
   End If
End Function
Public Function InitField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
   FieldName = FName
   FieldType = FType
   FieldValue = FValue
   FieldCat = FCat
End Function
