VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int32Formattable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IComparable
Implements IFormattable

Public mValue As Long


Private Sub Class_Terminate()
'    Debug.Print "Terminated"
End Sub

Private Function IComparable_CompareTo(Value As Variant) As Long
    Dim l As Long
    If IsObject(Value) Then
        Dim o As Int32
        Set o = Value
        l = o.Value
    Else
        l = Value
    End If
    
    If mValue < l Then IComparable_CompareTo = -1: Exit Function
    If mValue > l Then IComparable_CompareTo = 1
End Function

Private Function IFormattable_ToString(ByVal Format As String, ByVal Provider As VBCorLib.IFormatProvider) As String
    IFormattable_ToString = VBA.Format$(mValue, String$(Val(Format), "0"))
End Function

