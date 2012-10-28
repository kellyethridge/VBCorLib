VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IList"
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
'    Module: IList
'

''
' An interface that provides methods for accessing and manipulating
' values at specified indexes within a list.
'
' @remarks This interface should be implemented in classes that provide
' a list style functionality.
'
' @see ICollection
' @see IEnumerable
'
Option Explicit

''
' Returns the number of items in the list.
'
' @return The number of items in the list.
'
Public Property Get Count() As Long: End Property

''
' Copies the elements in the list to the array.
'
' @param arr The array to copy the elements to.
' @param Index The starting index in the array to begin copying to.
'
Public Sub CopyTo(ByRef arr As Variant, ByVal Index As Long): End Sub

''
' Returns an enumerator used to iterate over the list.
'
' @return An enumerator used to iterate over the list.
'
Public Function GetEnumerator() As IEnumerator: End Function

''
' Returns an IEnumVariant compatible enumerator used by For..Each.
'
' @return The enumerator that is used by For..Each enumeration.
'
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

''
' Returns if the list is fixed-size.
'
' @return Indication of the list being fixed-size.
' @remarks A fixed-size list can have the elements changed, however,
' elements cannot be added or removed.
'
Public Property Get IsFixedSize() As Boolean: End Property

''
' Returns if the list is read-only.
'
' @return Indication of the list being read-only.
' @remarks A read-only list can only view the elements in the list, however,
' elements cannot be added, removed, or modified.
'
Public Property Get IsReadOnly() As Boolean: End Property

''
' Returns an item at specified index.
'
' @param Index The index of the item to return.
' @return The item at the specified index.
'
Public Property Get Item(ByVal Index As Long) As Variant: End Property
Attribute Item.VB_UserMemId = 0

''
' Sets the item at the specified index to a new value.
'
' @param Index The index of the item to set the value to.
' @param RHS The new value to set the item to.
'
Public Property Let Item(ByVal Index As Long, ByRef RHS As Variant): End Property

''
' Sets the item at the specified index to a new value.
'
' @param Index The index of the item to set the value to.
' @param RHS The new value to set the item to.
'
Public Property Set Item(ByVal Index As Long, ByRef RHS As Variant): End Property

''
' Adds a new item to the list.
'
' @param value The item to be added to the list.
'
Public Function Add(ByRef value As Variant) As Long: End Function

''
' Clears the contents of the list.
'
Public Sub Clear(): End Sub

''
' Determines if the list contains a specific value.
'
' @param value the value to check the list for.
' @param comparer A custom comparer used for comparing values in the list to the value being sought.
' @return An indication of the value being found in the list.
'
Public Function Contains(ByRef value As Variant, Optional ByVal Comparer As IComparer) As Boolean: End Function

''
' Returns the index of the value within the list.
'
' @param value The value to find the index of in the list.
' @param comparer A custom comparer used for comparing values in the list to the value being sought.
' @return The index the value was found, or -1 if the value was not found.
' @remarks In general, -1 is returned for an unfound value. If the lowerbound of the underlying
' list is not 0, then the return value should be the lowerbound-1 for values not found.
'
Public Function IndexOf(ByRef value As Variant, Optional ByVal Comparer As IComparer) As Long: End Function

''
' Inserts a value into a list at the specified index.
'
' @param Index The postion in the list to insert the new value.
' @param value The new value to be inserted into the list.
'
Public Sub Insert(ByVal Index As Long, ByRef value As Variant): End Sub

''
' Removes a value from the list.
'
' @param value The value to be removed from the list.
' @param comparer A custom comparer used for comparing values in the list to the value being sought.
' @remarks This method searches the list for the first instance of the value. If one
' is found, it is removed from the list.
'
Public Sub Remove(ByRef value As Variant, Optional ByVal Comparer As IComparer): End Sub

''
' Removes a value from the list at the specified index.
'
' @param Index The position of the value in the list to be removed.
'
Public Sub RemoveAt(ByVal Index As Long): End Sub


