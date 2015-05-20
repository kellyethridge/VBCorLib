VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixedSizeArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: FixedSizeArrayList
'
Option Explicit
Implements IObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList As ArrayList


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As ArrayList)
    Set mList = List
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_BinarySearch = mList.BinarySearch(Value, Comparer)
End Function

Private Function ArrayList_BinarySearchEx(ByVal Index As Long, ByVal Count As Long, Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_BinarySearchEx = mList.BinarySearchEx(Index, Count, Value, Comparer)
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Function ArrayList_Clone() As ArrayList
    Dim Ret As New FixedSizeArrayList
    Ret.Init mList.Clone
    Set ArrayList_Clone = Ret
End Function

Private Function ArrayList_Contains(Value As Variant) As Boolean
    ArrayList_Contains = mList.Contains(Value)
End Function

Private Sub ArrayList_CopyTo(Arr As Variant, Optional ArrayIndex As Variant)
    mList.CopyTo Arr, ArrayIndex
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, Arr As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    mList.CopyToEx Index, Arr, ArrayIndex, Count
End Sub

Private Property Get ArrayList_Count() As Long
    ArrayList_Count = mList.Count
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Object
    Dim Range As ListRange
    Range = OptionalRange(Index, Count, 0, mList.Count)
    If Range.Index < 0 Then _
        Error.NegativeNumber "Index"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count > mList.Count Then _
        Error.InvalidOffsetLength
    
    Dim Ret As New ArrayListEnumerator
    Ret.Init Me, Range.Index, Range.Count
    Set ArrayList_GetEnumerator = Ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    If Index < 0 Then _
        Error.NegativeNumber "Index"
    If Count < 0 Then _
        Error.NegativeNumber "Count"
    If Index + Count > mList.Count Then _
        Error.InvalidOffsetLength
    
    Dim Ret As New RangedArrayList
    Ret.Init Me, Index, Count
    Set ArrayList_GetRange = Ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    ArrayList_IndexOf = mList.IndexOf(Value, StartIndex, Count)
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, Source As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = True
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = False
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, Value As Variant)
    Set mList(Index) = Value
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, Value As Variant)
    mList(Index) = Value
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant ArrayList_Item, mList(Index)
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    ArrayList_LastIndexOf = mList.LastIndexOf(Value, StartIndex, Count)
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    mList.Reverse Index, Count
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, Source As Variant)
    mList.SetRange Index, Source
End Sub

Private Sub ArrayList_Sort(Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    mList.Sort Index, Count, Comparer
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As VbVarType = 12&) As Variant
    ArrayList_ToArray = mList.ToArray(ArrayType)
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = mList.ToString
End Function

Private Sub ArrayList_TrimToSize()
    Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_FixedSizeCollection))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = ArrayList_Clone
End Function


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
    ArrayList_CopyTo Arr, Index
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = ArrayList_Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(Value As Variant) As Long
    IList_Add = ArrayList_Add(Value)
End Function

Private Sub IList_Clear()
    ArrayList_Clear
End Sub

Private Function IList_Contains(Value As Variant) As Boolean
    IList_Contains = ArrayList_Contains(Value)
End Function

Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long)
    ArrayList_CopyTo Arr, Index
End Sub

Private Property Get IList_Count() As Long
    IList_Count = ArrayList_Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IList_IndexOf(Value As Variant) As Long
    IList_IndexOf = ArrayList_IndexOf(Value)
End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant)
    ArrayList_Insert Index, Value
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = ArrayList_IsFixedSize
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = ArrayList_IsReadOnly
End Property

Private Property Set IList_Item(ByVal Index As Long, Value As Variant)
    Set ArrayList_Item(Index) = Value
End Property

Private Property Let IList_Item(ByVal Index As Long, Value As Variant)
    ArrayList_Item(Index) = Value
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant IList_Item, ArrayList_Item(Index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = ArrayList_NewEnum
End Function

Private Sub IList_Remove(Value As Variant)
    ArrayList_Remove Value
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    ArrayList_RemoveAt Index
End Sub

