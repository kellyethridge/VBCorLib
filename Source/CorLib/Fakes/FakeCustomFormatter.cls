VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeCustomFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IFormatProvider
Implements ICustomFormatter

Public ReturnNull As Boolean
Public FormatResult As String
Public FormatText As String
Public FormatType As String
Public FormatArg As Variant
Public Disable As Boolean


Private Function ICustomFormatter_Format(ByVal FormatText As String, Arg As Variant, ByVal Provider As VBCorLib.IFormatProvider) As String
    Me.FormatText = FormatText
    Me.FormatArg = Arg
    
    If ReturnNull Then
        ICustomFormatter_Format = vbNullString
    Else
        ICustomFormatter_Format = FormatResult
    End If
End Function

Private Function IFormatProvider_GetFormat(ByRef FormatType As String) As Object
    Me.FormatType = FormatType
    If Not Me.Disable And LCase$(FormatType) = "icustomformatter" Then
        Set IFormatProvider_GetFormat = Me
    End If
End Function
