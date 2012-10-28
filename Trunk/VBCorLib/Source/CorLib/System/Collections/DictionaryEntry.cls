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
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: DictionaryEntry
'

''
' A DictionaryEntry is used to contain a key/value pair. This enables collections
' to keep the two values together throughout an operation.
'
' @see Constructors
'
Option Explicit
Implements IObject

Private Const PROP_KEY As String = "Key"
Private Const PROP_VALUE As String = "Value"


Private mKey    As Variant
Private mValue  As Variant


''
' Returns the Key protion of the entry.
'
' @return The key from the entry.
'
Public Property Get Key() As Variant
    Call VariantCopy(Key, mKey)
End Property

''
' Sets the key portion of the entry.
'
' @param RHS The new key to set the entry to.
'
Public Property Let Key(ByRef RHS As Variant)
    mKey = RHS
End Property

''
' Sets the key portion of the entry.
'
' @param RHS The new key to set the entry to.
'
Public Property Set Key(ByRef RHS As Variant)
    Set mKey = RHS
End Property

''
' Returns the value portion of the entry pair.
'
' @return The value from the entry.
'
Public Property Get Value() As Variant
    Call VariantCopy(Value, mValue)
End Property

''
' Sets the value portion of the entry.
'
' @param RHS The new value to set the entry to.
'
Public Property Let Value(ByRef RHS As Variant)
    mValue = RHS
End Property

''
' Sets the value portion of the entry.
'
' @param RHS The new value to set the entry to.
'
Public Property Set Value(ByRef RHS As Variant)
    Set mValue = RHS
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Key As Variant, ByVal Value As Variant)
    Call Helper.MoveVariant(mKey, Key)
    Call Helper.MoveVariant(mValue, Value)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Call Helper.MoveVariant(mKey, .ReadProperty(PROP_KEY, Empty))
        Call Helper.MoveVariant(mValue, .ReadProperty(PROP_VALUE, Empty))
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_KEY, mKey)
        Call .WriteProperty(PROP_VALUE, mValue)
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
