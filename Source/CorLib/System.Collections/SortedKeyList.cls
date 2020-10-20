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
' Module: SortedKeyList
'
'@Folder("CorLib.System.Collections")
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
'   ICollection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long)
    mList.CopyKeysTo Arr, Index
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
'   IEnumerable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = mList.GetKeyEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(Value As Variant) As Long
    Error.NotSupported NotSupported_ReadOnlyCollection
End Function

Private Sub IList_Clear()
    Error.NotSupported NotSupported_ReadOnlyCollection
End Sub

Private Function IList_Contains(Value As Variant) As Boolean
    IList_Contains = mList.Contains(Value)
End Function

Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long)
    mList.CopyKeysTo Arr, Index
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
    Error.NotSupported NotSupported_ReadOnlyCollection
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = True
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = True
End Property

Private Property Set IList_Item(ByVal Index As Long, Value As Variant)
    Error.NotSupported NotSupported_ReadOnlyCollection
End Property

Private Property Let IList_Item(ByVal Index As Long, Value As Variant)
    Error.NotSupported NotSupported_ReadOnlyCollection
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant IList_Item, mList.GetKey(Index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = NewEnum
End Function

Private Sub IList_Remove(Value As Variant)
    Error.NotSupported NotSupported_ReadOnlyCollection
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    Error.NotSupported NotSupported_ReadOnlyCollection
End Sub
