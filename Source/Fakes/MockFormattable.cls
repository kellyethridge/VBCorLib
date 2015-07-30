VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockFormattable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IFormattable

Public Result As String
Public Format As String
Public Provider As IFormatProvider

Private Function IFormattable_ToString(ByRef Format As String, ByVal Provider As IFormatProvider) As String
    Me.Format = Format
    Set Me.Provider = Provider
    IFormattable_ToString = Me.Result
End Function
