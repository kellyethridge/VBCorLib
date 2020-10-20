VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: DictionaryEntry
'

''
' A DictionaryEntry is used to contain a key/value pair. This enables collections
' to keep the two values together throughout an operation.
'
' @see Constructors
'
'@Folder("CorLib.System.Collections")
Option Explicit
Implements IObject

Private Const PropKey      As String = "Key"
Private Const PropValue    As String = "Value"

Private mKey    As Variant
Private mValue  As Variant


''
' Returns the Key protion of the entry.
'
' @return The key from the entry.
'
Public Property Get Key() As Variant
    VariantCopy Key, mKey
End Property

''
' Sets the key portion of the entry.
'
' @param Value The new key to set the entry to.
'
Public Property Let Key(ByRef Value As Variant)
    mKey = Value
End Property

''
' Sets the key portion of the entry.
'
' @param Value The new key to set the entry to.
'
Public Property Set Key(ByRef Value As Variant)
    Set mKey = Value
End Property

''
' Returns the value portion of the entry pair.
'
' @return The value from the entry.
'
Public Property Get Value() As Variant
    VariantCopy Value, mValue
End Property

''
' Sets the value portion of the entry.
'
' @param Value The new value to set the entry to.
'
Public Property Let Value(ByRef Value As Variant)
    mValue = Value
End Property

''
' Sets the value portion of the entry.
'
' @param Value The new value to set the entry to.
'
Public Property Set Value(ByRef Value As Variant)
    Set mValue = Value
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As DictionaryEntry
    
    On Error GoTo Catch
    Set Other = Value
    
    If Not Other Is Nothing Then
        Equals = Object.Equals(mKey, Other.Key) And Object.Equals(mValue, Other.Value)
    End If
    
Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = Object.GetHashCode(mKey) Xor Object.GetHashCode(mValue)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Key As Variant, ByRef Value As Variant)
    VariantCopyInd mKey, Key
    VariantCopyInd mValue, Value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Helper.MoveVariant mKey, .ReadProperty(PropKey, Empty)
        Helper.MoveVariant mValue, .ReadProperty(PropValue, Empty)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropKey, mKey
        .WriteProperty PropValue, mValue
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
