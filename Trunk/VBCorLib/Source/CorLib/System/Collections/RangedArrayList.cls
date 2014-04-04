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
    mVersion = List.Version
    mIndex = Index
    mCount = Count
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyVersion()
    If mVersion <> mList.Version Then _
        Throw Cor.NewInvalidOperationException("The base ArrayList has been modified outside of the Ranged list.")
End Sub

Private Function GetCollectionSize(ByRef c As Variant) As Long
    If IsArray(c) Then
        If cArray.IsNull(c) Then _
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
    Call VerifyVersion
    Call mList.Insert(mIndex + mCount, Value)
    ArrayList_Add = mCount
    mCount = mCount + 1
    mVersion = mVersion + 1
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Call VerifyVersion
    Call mList.InsertRange(mIndex + mCount, c)
    mCount = mCount + GetCollectionSize(c)
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mCount, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim i As Long
    i = mList.BinarySearch(Value, mIndex + ElemIndex, ElemCount, Comparer)
    If i >= 0 Then i = i - mIndex
    ArrayList_BinarySearch = i
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    Call VerifyVersion
    mList.Capacity = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Capacity() As Long
    Call VerifyVersion
    ArrayList_Capacity = mList.Capacity
End Property

Private Sub ArrayList_Clear()
    Call VerifyVersion
    If mCount = 0 Then Exit Sub
    Call mList.RemoveRange(mIndex, mCount)
    mCount = 0
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_Clone() As ArrayList
    Call VerifyVersion
    Dim Ret As New RangedArrayList
    Call Ret.Init(mList.Clone, mIndex, mCount)
    Set ArrayList_Clone = Ret
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    Call VerifyVersion
    ArrayList_Contains = (mList.IndexOf(Value, mIndex, mCount, Comparer) >= 0)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional ArrayIndex As Variant)
    Call VerifyVersion
    
    Dim Index As Long
    If IsMissing(ArrayIndex) Then Index = LBound(DstArray) Else Index = ArrayIndex
    Call mList.CopyToEx(mIndex, DstArray, Index, mCount)
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    Call VerifyVersion
    
    Dim Result As Long
    Result = VerifyListRange(mCount, Index, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, Index, "Index", Count, "Count")
    
    Call mList.CopyToEx(mIndex + Index, DstArray, ArrayIndex, Count)
End Sub

Private Property Get ArrayList_Count() As Long
    Call VerifyVersion
    ArrayList_Count = mCount
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Object
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mCount, StartIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))
    
    
    Dim Ret As New ArrayListEnumerator
    Call Ret.Init(Me, ElemIndex, ElemCount)
    Set ArrayList_GetEnumerator = Ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    Dim Result As Long
    Result = VerifyListRange(mCount, Index, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, Index, "Index", Count, "Count")
    
    Dim Ret As RangedArrayList
    Set Ret = New RangedArrayList
    Call Ret.Init(Me, Index, Count)
    Set ArrayList_GetRange = Ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mCount, StartIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))
    
    Dim Index As Long
    Index = mList.IndexOf(Value, mIndex + ElemIndex, ElemCount, Comparer)
    If Index >= 0 Then Index = Index - mIndex
    ArrayList_IndexOf = Index
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    Call VerifyVersion
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Call mList.Insert(mIndex + Index, Value)
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    Call VerifyVersion
    
    If Index < 0 Or Index > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Dim TotalElements As Long
    TotalElements = GetCollectionSize(c)
    Call mList.InsertRange(mIndex + Index, c)
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
    Call VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Set mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Call VerifyVersion
    
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    mList(mIndex + Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Call VerifyVersion
    
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Helper.MoveVariant ArrayList_Item, mList(mIndex + Index)
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Call VerifyVersion
    
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
    Index = mList.LastIndexOf(Value, mIndex + ElemIndex, ElemCount, Comparer)
    If Index >= 0 Then Index = Index - mIndex
    ArrayList_LastIndexOf = Index
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    Call VerifyVersion
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant, Optional ByVal Comparer As IComparer)
    Call VerifyVersion
    
    Dim Index As Long
    Index = mList.IndexOf(Value, mIndex, mCount, Comparer)
    If Index >= 0 Then
        Call mList.RemoveAt(Index)
        mCount = mCount - 1
        mVersion = mVersion + 1
    End If
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Call VerifyVersion
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Call mList.RemoveAt(mIndex + Index)
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Call VerifyVersion
    
    If Index < 0 Or Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(IIf(Index < 0, "Index", "Count"), Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidOffLen), "count")
    
    Call mList.RemoveRange(mIndex + Index, Count)
    mCount = mCount - Count
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mCount, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Call mList.Reverse(mIndex + ElemIndex, ElemCount)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    Call VerifyVersion
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
    Dim TotalElements As Long
    TotalElements = GetCollectionSize(c)
    If Index + TotalElements > mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_LargerThanCollection))
    
    Call mList.SetRange(mIndex + Index, c)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Sort(Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    Call VerifyVersion
    
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mCount, StartIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(StartIndex))

    Call mList.Sort(mIndex + ElemIndex, ElemCount, Comparer)
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As vbVarType = 12&) As Variant
    Call VerifyVersion
    
    Dim Ret As Variant
    Ret = cArray.CreateInstance(ArrayType, mCount)
    If mCount > 0 Then Call mList.CopyToEx(mIndex, Ret, 0, mCount)
    ArrayList_ToArray = Ret
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = IObject_ToString
End Function

Private Sub ArrayList_TrimToSize()
    Call VerifyVersion
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
    IObject_Equals = Object.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = ObjPtr(CUnk(Me))
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long)
    Call ArrayList_CopyTo(Arr, Index)
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

Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    IList_Contains = ArrayList_Contains(Value, Comparer)
End Function

Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long)
    Call ArrayList_CopyTo(Arr, Index)
End Sub

Private Property Get IList_Count() As Long
    IList_Count = ArrayList_Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = ArrayList_GetEnumerator
End Function

Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long
    IList_IndexOf = ArrayList_IndexOf(Value, , , Comparer)
End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant)
    Call ArrayList_Insert(Index, Value)
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = ArrayList_IsFixedSize
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = ArrayList_IsReadOnly
End Property

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant)
    Set ArrayList_Item(Index) = RHS
End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant)
    ArrayList_Item(Index) = RHS
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Call Helper.MoveVariant(IList_Item, ArrayList_Item(Index))
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = ArrayList_NewEnum
End Function

Private Sub IList_Remove(Value As Variant, Optional ByVal Comparer As IComparer)
    Call ArrayList_Remove(Value, Comparer)
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    Call ArrayList_RemoveAt(Index)
End Sub


