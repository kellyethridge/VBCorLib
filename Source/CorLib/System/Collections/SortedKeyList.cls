VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedKeyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: SortedKeyList
'
Option Explicit
Implements IObject
Implements IList
Implements ICollection
Implements IEnumerable

Private mList As SortedList



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = CreateEnumerator(mList.GetKeyEnumerator)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As SortedList)
    Set mList = List
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long)
    Call mList.CopyKeysTo(Arr, Index)
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = mList.Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = mList.GetKeyEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = NewEnum
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = mList.GetKeyEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(Value As Variant) As Long
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Function

Private Sub IList_Clear()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Function IList_Contains(Value As Variant) As Boolean
    IList_Contains = mList.Contains(Value)
End Function

Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long)
    Call mList.CopyKeysTo(Arr, Index)
End Sub

Private Property Get IList_Count() As Long
    IList_Count = mList.Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = mList.GetKeyEnumerator
End Function

Private Function IList_IndexOf(Value As Variant) As Long
    IList_IndexOf = mList.IndexOfKey(Value)
End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = True
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = True
End Property

Private Property Set IList_Item(ByVal Index As Long, Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Let IList_Item(ByVal Index As Long, Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Call Helper.MoveVariant(IList_Item, mList.GetKey(Index))
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = NewEnum
End Function

Private Sub IList_Remove(Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_ReadOnlyCollection))
End Sub
