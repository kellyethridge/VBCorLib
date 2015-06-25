VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeFormatProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IFormatProvider

Public GetFormatResult  As Object
Public FormatType       As String


Private Function IFormatProvider_GetFormat(FormatType As String) As Object
    Me.FormatType = FormatType
    Set IFormatProvider_GetFormat = GetFormatResult
End Function
