VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RangedArrayList"
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
' Module: RangedArrayList
'

''
' This class is used to represent a sub-array view of the parent ArrayList object.
'

Option Explicit
Implements IObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList       As ArrayList
Private mVersion    As Long
Private mIndex      As Long
Private mCount      As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As ArrayList, ByVal Index As Long, ByVal Count As Long)
    Set mList = List
    mVersion = GetCollectionVersion(List)
    mIndex = Index
    mCount = Count
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyVersion()
    If mVersion <> GetCollectionVersion(mList) Then
        Throw Cor.NewInvalidOperationException("The base ArrayList has been modified outside of the Ranged list.")
    End If
End Sub

Private Function GetCollectionSize(ByRef c As Variant) As Long
    If IsArray(c) Then
        If CorArray.IsNull(c) Then _
            Throw Cor.NewArgumentNullException("c", Environment.GetResourceString(ArgumentNull_Array))
        
        GetCollectionSize = UBound(c) - LBound(c) + 1
    
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw Cor.NewArgumentNullException("c", Environment.GetResourceString(ArgumentNull_Collection))
        
        If TypeOf c Is Collection Then
            Dim Col As Collection
            Set Col = c
            GetCollectionSize = Col.Count
        ElseIf TypeOf c Is ICollection Then
            Dim ICol As ICollection
            Set ICol = c
            GetCollectionSize = ICol.Count
        Else
            Throw Cor.NewInvalidCastException("An ICollection or VBA.Collection object is required.")
        End If
    Else
        Throw Cor.NewInvalidCastException("An ICollection object, VBA.Collection or an Array is required.")
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    VerifyVersion
    mList.Insert mIndex + mCount, Value
    ArrayList_Add = mCount
    mCount = mCount + 1
    mVersion = mVersion + 1
End Function

Private Sub ArrayList_AddRange(c As Variant)
    VerifyVersion
    mList.InsertRange mIndex + mCount, c
    mCount = mCount + GetCollectionSize(c)
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByVal Comparer As IComparer) As Long
    VerifyVersion
    ArrayList_BinarySearch = ArrayList_BinarySearchEx(0, mCount, Value, Comparer)
End Function

Private Function ArrayList_BinarySearchEx(ByVal Index As Long, ByVal Count As Long, Value As Variant, Optional ByVal Comparer As IComparer) As Long
    VerifyVersion
    
    If Index < 0 Then _
        Error.NegativeNumber "Index"
    If Count < 0 Then _
        Error.NegativeNumber "Count"
    If Index + Count > mCount Then _
        Error.InvalidOffsetLength
    
    Dim ItemIndex As Long
    ItemIndex = mList.BinarySearchEx(mIndex + Index, Count, Value, Comparer)
    If ItemIndex >= 0 Then
        ItemIndex = ItemIndex - mIndex
    End If
    
    ArrayList_BinarySearchEx = ItemIndex
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    VerifyVersion
    mList.Capacity = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Capacity() As Long
    VerifyVersion
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    VerifyVersion
    If mCount > 0 Then
        mList.RemoveRange mIndex, mCount
        mCount = 0
        mVersion = mVersion + 1
    End If
End Sub

Private Function ArrayList_Clone() As ArrayList
    VerifyVersion
    Dim Ret As New RangedArrayList
    Ret.Init mList.Clone, mIndex, mCount
    Set ArrayList_Clone = Ret
End Function

Private Function ArrayList_Contains(Value As Variant) As Boolean
    VerifyVersion
    ArrayList_Contains = (mList.IndexOf(Value, mIndex, mCount) >= 0)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional ArrayIndex As Variant)
    VerifyVersion
    
    Dim Index As Long
    If IsMissing(ArrayIndex) Then
        Index = LBound(DstArray)
    Else
        Index = ArrayIndex
    End If
    
    mList.CopyToEx mIndex, DstArray, Index, mCount
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    VerifyVersion
    
    Dim Result As Long
    Result = VerifyArrayRange(GetArrayPointer(DstArray), ArrayIndex, Count)
    If Result <> NO_ERROR Then _
        ThrowArrayRangeException Result, "DstArray", ArrayIndex, "ArrayIndex", Count, "Count"
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Count, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    mList.CopyToEx mIndex + Index, DstArray, ArrayIndex, Count
End Sub

Private Property Get ArrayList_Count() As Long
    VerifyVersion
    ArrayList_Count = mCount
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Object
    VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    ElemIndex = GetOptionalLong(Index, 0)
    ElemCount = GetOptionalLong(Count, mCount - ElemIndex)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", ElemIndex, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    Dim Ret As New ArrayListEnumerator
    Ret.Init Me, ElemIndex, ElemCount
    Set ArrayList_GetEnumerator = Ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Count, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    Dim Ret As RangedArrayList
    Set Ret = New RangedArrayList
    Ret.Init Me, Index, Count
    Set ArrayList_GetRange = Ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    ElemIndex = GetOptionalLong(StartIndex, 0)
    ElemCount = GetOptionalLong(Count, mCount)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", ElemIndex, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
           
    Dim Result As Long
    Result = mList.IndexOf(Value, mIndex + ElemIndex, ElemCount)
    If Result >= 0 Then
        Result = Result - mIndex
    End If
    
    ArrayList_IndexOf = Result
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    VerifyVersion
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    mList.Insert mIndex + Index, Value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    VerifyVersion
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Dim TotalElements As Long
    TotalElements = GetCollectionSize(c)
    mList.InsertRange mIndex + Index, c
    mCount = mCount + TotalElements
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = mList.IsFixedSize
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = mList.IsReadOnly
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Set mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    VerifyVersion
    
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    VerifyVersion
    
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Helper.MoveVariant ArrayList_Item, mList(mIndex + Index)
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    VerifyVersion
    
    Dim ElemIndex As Long
    Dim ElemCount As Long
    If GetOptionalLongPair(StartIndex, mCount - 1, ElemIndex, Count, mCount, ElemCount) = Argument_ParamRequired Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), IIf(IsMissing(StartIndex), "StartIndex", "Count"))
    If ElemIndex >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", Message:=Environment.GetResourceString(ArgumentOutOfRange_UBound))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex - ElemCount + 1 < 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen), "Count")

    Dim Index As Long
    Index = mList.LastIndexOf(Value, mIndex + ElemIndex, ElemCount)
    If Index >= 0 Then
        Index = Index - mIndex
    End If
    
    ArrayList_LastIndexOf = Index
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    VerifyVersion
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant)
    VerifyVersion
    
    Dim Index As Long
    Index = mList.IndexOf(Value, mIndex, mCount)
    If Index >= 0 Then
        mList.RemoveAt Index
        mCount = mCount - 1
        mVersion = mVersion + 1
    End If
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    mList.RemoveAt mIndex + Index
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    VerifyVersion
    
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(IIf(Index < 0, "Index", "Count"), Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen), "count")
    
    mList.RemoveRange mIndex + Index, Count
    mCount = mCount - Count
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    VerifyVersion
    
    Dim ElemCount As Long
    Dim ElemIndex As Long
    
    If GetOptionalLongPair(Index, 0, ElemIndex, Count, mCount, ElemCount) = Argument_ParamRequired Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), IIf(IsMissing(Index), "Index", "Count"))
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", ElemIndex, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))
    
    mList.Reverse mIndex + ElemIndex, ElemCount
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    VerifyVersion
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
    Dim TotalElements As Long
    TotalElements = GetCollectionSize(c)
    If Index + TotalElements > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_LargerThanCollection))
    
    mList.SetRange mIndex + Index, c
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Sort(Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    VerifyVersion
    
    Dim ElemCount As Long
    Dim ElemIndex As Long
    ElemIndex = GetOptionalLong(Index, 0)
    ElemCount = GetOptionalLong(Count, mCount - ElemIndex)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", ElemIndex, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen))

    mList.Sort mIndex + ElemIndex, ElemCount, Comparer
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As VbVarType = 12&) As Variant
    VerifyVersion
    
    ArrayList_ToArray = CorArray.CreateInstance(ArrayType, mCount)
    If mCount > 0 Then
        mList.CopyToEx mIndex, ArrayList_ToArray, 0, mCount
    End If
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = IObject_ToString
End Function

Private Sub ArrayList_TrimToSize()
    VerifyVersion
    mList.TrimToSize
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_Version() As Long
    ArrayList_Version = mVersion
End Property


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


