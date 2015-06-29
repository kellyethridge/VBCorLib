VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyGuid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: ReadOnlyGuid
'
Option Explicit
Implements IObject
Implements Guid
Implements IComparable
Implements IFormattable


Private mGuid As Guid


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal g As Guid)
    Set mGuid = g
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = mGuid.Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = mGuid.GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = mGuid.ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Guid Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Guid_CompareTo(Value As Variant) As Long
    Guid_CompareTo = mGuid.CompareTo(Value)
End Function

Private Function Guid_Equals(Value As Variant) As Boolean
    Guid_Equals = mGuid.Equals(Value)
End Function

Private Function Guid_GetHashCode() As Long
    Guid_GetHashCode = mGuid.GetHashCode
End Function

Private Property Get Guid_Handle() As Long
    Throw Cor.NewNotSupportedException("Guid is read-only.")
End Property

Private Property Get Guid_IsReadOnly() As Boolean
    Guid_IsReadOnly = True
End Property

Private Function Guid_ToByteArray() As Byte()
    Guid_ToByteArray = mGuid.ToByteArray
End Function

Private Function Guid_ToString(Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Guid_ToString = mGuid.ToString(Format, Provider)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = mGuid.CompareTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(ByRef Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = mGuid.ToString(Format, Provider)
End Function
