VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyArrayList"
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
' Module: ReadOnlyArrayList
'

''
' Wraps an ArrayList object in a read-only wrapper.

Option Explicit
Implements IObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList
Implements IVersionableCollection

Private mList As ArrayList


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As ArrayList)
    Set mList = List
End Sub

Friend Sub InitClone(ByVal List As ArrayList)
    Init List.Clone
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ThrowNotSupported()
    Error.NotSupported NotSupported_ReadOnlyCollection
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    ThrowNotSupported
End Function

Private Sub ArrayList_AddRange(c As Variant)
    ThrowNotSupported
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_BinarySearch = mList.BinarySearch(Value, Comparer)
End Function

Private Function ArrayList_BinarySearchEx(ByVal Index As Long, ByVal Count As Long, Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_BinarySearchEx = mList.BinarySearchEx(Index, Count, Value, Comparer)
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    ThrowNotSupported
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    ThrowNotSupported
End Sub

Private Function ArrayList_Clone() As ArrayList
    Dim Clone As New ReadOnlyArrayList
    
    Clone.InitClone mList
    Set ArrayList_Clone = Clone
End Function

Private Function ArrayList_Contains(Value As Variant) As Boolean
    ArrayList_Contains = mList.Contains(Value)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional ArrayIndex As Variant)
    mList.CopyTo DstArray, ArrayIndex
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    mList.CopyToEx Index, DstArray, ArrayIndex, Count
End Sub

Private Property Get ArrayList_Count() As Long
    ArrayList_Count = mList.Count
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef Index As Variant, Optional ByRef Count As Variant) As IEnumerator
    Dim Range       As ListRange
    Dim Enumerator  As New ArrayListEnumerator

    Range = MakeDefaultRange(Index, 0, Count, mList.Count)

    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_NeedNonNegNum
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > mList.Count Then _
        Error.Argument Argument_InvalidOffLen

    Enumerator.Init Me, Range.Index, Range.Count
    Set ArrayList_GetEnumerator = Enumerator
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashCode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    Dim List As New RangedArrayList

    If Index < 0 Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_NeedNonNegNum
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Index + Count > mList.Count Then _
        Error.Argument Argument_InvalidOffLen
    
    List.Init Me, Index, Count
    Set ArrayList_GetRange = List
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    ArrayList_IndexOf = mList.IndexOf(Value, StartIndex, Count)
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    ThrowNotSupported
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    ThrowNotSupported
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = True
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = True
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    ThrowNotSupported
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    ThrowNotSupported
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
    ThrowNotSupported
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    ThrowNotSupported
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    ThrowNotSupported
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    ThrowNotSupported
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    ThrowNotSupported
End Sub

Private Sub ArrayList_Sort(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    ThrowNotSupported
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As VbVarType = 12&) As Variant
    ArrayList_ToArray = mList.ToArray(ArrayType)
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = mList.ToString
End Function

Private Sub ArrayList_TrimToSize()
    ThrowNotSupported
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = ArrayList_Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection
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
'   IEnumerable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = ArrayList_NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IVersionableCollection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IVersionableCollection_CurrentVersion() As Long
    IVersionableCollection_CurrentVersion = GetCollectionVersion(mList)
End Property
