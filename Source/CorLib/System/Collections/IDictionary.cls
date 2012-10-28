VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: IDictionary
'

''
' This provides a standard interface to be impelments by key-value list collection.
'
' @remarks<br>
' Lists that maintain a key-value pair should implement this interface to allow
' cross-compatibility usage amongst all key-value pair lists.
'
Option Explicit

''
' Returns an IEnumerator object.
'
' @return The enumerator.
'
Public Function GetEnumerator() As IEnumerator: End Function

''
' Returns an IEnumVariant object.
'
' @return The enumerator
' @remarks The returned object is an IEnumVariant compatible wrapper
' that delegates the correct calls to the IEnumerator object returned
' by GetEnumerator.
'
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

''
' Returns the size of the collection.
'
' @return The size of the collection.
'
Public Property Get Count() As Long: End Property

''
' Copies the contents of the collection to an array.
'
' @param arr The array to copy the elements into.
' @param index The index in the destination array to start copying the elements to.
'
Public Sub CopyTo(ByRef arr As Variant, ByVal Index As Long): End Sub

''
' Returns if this instance is fixed-size.
'
' @return Indication if is a fixid-size dictionary list.
' @remarks A FixedSize list can still have items replaced by new values, only now
' values can be added or removed.
'
Public Property Get IsFixedSize() As Boolean: End Property

''
' Returns if this instance is read-only.
'
' @return An indication if this instance is read-only.
' @remarks A ReadOnly list can only retrieve items from the list. The list
' cannot change, add, or remove items.
'
Public Property Get IsReadOnly() As Boolean: End Property

''
' Returns an item from the list that is associated with the specified key.
'
' @param Key The key to search for in the list.
' @return The value that is associated to the key.
'
Public Property Get Item(ByRef Key As Variant) As Variant: End Property
Attribute Item.VB_UserMemId = 0

''
' Sets the value in the list that is associated to the specified key.
'
' @param Key The key used to identify the value.
' @param value The value to associate with the key.
'
Public Property Let Item(ByRef Key As Variant, ByRef value As Variant): End Property

''
' Sets the value in the list that is associated to the specified key.
'
' @param Key The key used to identify the value.
' @param value The value to associate with the key.
'
Public Property Set Item(ByRef Key As Variant, ByRef value As Variant): End Property

''
' Returns a collection of the keys in the list.
'
' @return A collection of the keys in the list.
'
Public Property Get Keys() As ICollection: End Property

''
' Returns a collection of the values in the list.
'
' @return A collection of the values in the list.


Public Property Get values() As ICollection: End Property

''
' Adds a new value that is associated with the key to the list.
'
' @param Key The new key used to locate the value.
' @param value The value to be associated with the key.
'
Public Sub Add(ByRef Key As Variant, ByRef value As Variant): End Sub

''
' Clears the list of its contents.
'
Public Sub Clear(): End Sub

''
' Searches for the existence of the key in the list.
'
' @param Key The key to be searching for.
' @return Indication of the existence of the key.
'
Public Function Contains(ByRef Key As Variant) As Boolean: End Function

''
' Removes a value from the list that is associated to the key.
'
' @param Key The key to search for to be removed.
'
Public Sub Remove(ByRef Key As Variant): End Sub
