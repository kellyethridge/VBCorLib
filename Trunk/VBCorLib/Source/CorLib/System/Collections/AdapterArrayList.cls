VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdapterArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2008 Kelly Ethridge
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
'    Module: AdapterArrayList
'

Option Explicit
Implements IObject
Implements ArrayList
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList

Private mList       As IList
Private mComparer   As IComparer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal list As IList)
    Set mList = list
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub QuickSortList(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: Call Helper.MoveVariant(x, mList((i + j) \ 2))
        Do
            Do While mComparer.Compare(mList(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Compare(mList(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then
                Dim t1 As Variant
                Dim t2 As Variant
                
                Call Helper.MoveVariant(t1, mList(i))
                Call Helper.MoveVariant(t2, mList(j))
                
                If IsObject(t1) Then
                    Set mList(j) = t1
                Else
                    mList(j) = t1
                End If
                
                If IsObject(t2) Then
                    Set mList(i) = t2
                Else
                    mList(i) = t2
                End If
            End If
                
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call QuickSortList(Left, j)
            Left = i
        Else
            If i < Right Then Call QuickSortList(i, Right)
            Right = j
        End If
    Loop
End Sub

Private Sub InsertItems(ByVal Index As Long, c As Variant)
    Dim v As Variant
    
    For Each v In c
        Call mList.Insert(Index, v)
        Index = Index + 1
    Next v
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    ArrayList_Add = mList.Add(Value)
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Dim v As Variant
    
    For Each v In c
        Call mList.Add(v)
    Next v
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByRef startIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mList.Count, startIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "startIndex", ElemCount, "Count", IsMissing(startIndex))
    
    Dim MiddleIndex As Long
    Dim LowIndex    As Long
    Dim HighIndex   As Long
    
    If Comparer Is Nothing Then
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    LowIndex = ElemIndex
    HighIndex = LowIndex + ElemCount - 1
    Do While LowIndex <= HighIndex
        MiddleIndex = (LowIndex + HighIndex) \ 2
        Select Case Comparer.Compare(mList(MiddleIndex), Value)
            Case 0
                ArrayList_BinarySearch = MiddleIndex
                Exit Function
            Case Is > 0
                HighIndex = MiddleIndex - 1
            Case Else
                LowIndex = MiddleIndex + 1
        End Select
    Loop
    
    ArrayList_BinarySearch = Not LowIndex
End Function

Private Property Let ArrayList_Capacity(ByVal RHS As Long)
    If RHS < mList.Count Then
        Throw Cor.NewArgumentOutOfRangeException("Capacity cannot be set less than the size of the list.")
    End If
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.Count
End Property

Private Sub ArrayList_Clear()
    Call mList.Clear
End Sub

Private Function ArrayList_Clone() As ArrayList
    Dim Ret As New AdapterArrayList
    Call Ret.Init(mList)
    Set ArrayList_Clone = Ret
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    ArrayList_Contains = mList.Contains(Value, Comparer)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional ArrayIndex As Variant)
    Dim Index As Long
    Index = GetOptionalLong(ArrayIndex, 0)
    
    Call mList.CopyTo(DstArray, Index)
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    Dim Result As Long
    Result = VerifyListRange(mList.Count, Index, Count)
    If Result Then Call ThrowListRangeException(Result, Index, "Index", Count, "Count")
    Result = VerifyArrayRange(GetArrayPointer(DstArray, False), ArrayIndex, Count)
    If Result Then Call ThrowArrayRangeException(Result, "DstArray", ArrayIndex, "ArrayIndex", Count, "Count")
    
    Dim v As Variant
    Dim i As Long
    
    For i = 0 To Count - 1
        Call Helper.MoveVariant(v, mList(Index + i))
        If IsObject(v) Then
            Set DstArray(ArrayIndex + i) = v
        Else
            DstArray(ArrayIndex + i) = v
        End If
    Next i
End Sub

Private Property Get ArrayList_Count() As Long
    ArrayList_Count = mList.Count
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = IObject_Equals(Value)
End Function

Private Function ArrayList_GetEnumerator(Optional ByRef startIndex As Variant, Optional ByRef Count As Variant) As Object
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mList.Count, startIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(startIndex))
    
    Dim Ret As New ArrayListEnumerator
    Call Ret.Init(Me, ElemIndex, ElemCount)
    Set ArrayList_GetEnumerator = Ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    Dim Result As Long
    Result = VerifyListRange(mList.Count, Index, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, Index, "Index", Count, "Count")
    
    Dim Ret As RangedArrayList
    Set Ret = New RangedArrayList
    Call Ret.Init(Me, Index, Count)
    Set ArrayList_GetRange = Ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef startIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mList.Count, startIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(startIndex))
    
    If Comparer Is Nothing Then
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    Dim i As Long
    For i = ElemIndex To ElemIndex + ElemCount - 1
        If Comparer.Compare(mList(i), Value) = 0 Then
            ArrayList_IndexOf = i
            Exit Function
        End If
    Next i
    
    ArrayList_IndexOf = -1
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    Call mList.Insert(Index, Value)
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    If IsArray(c) Then
        Call InsertItems(Index, c)
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw New ArgumentNullException
            
        If TypeOf c Is ICollection Then
            Call InsertItems(Index, c)
        ElseIf TypeOf c Is IEnumerable Then
            Call InsertItems(Index, c)
        ElseIf TypeOf c Is Collection Then
            Call InsertItems(Index, c)
        Else
            Throw Cor.NewArgumentException("Only arrays, ICollection, Collection, and IEnumerable objects are supported.")
        End If
    Else
        Throw Cor.NewArgumentException("Only arrays, ICollection, Collection, and IEnumerable objects are supported.")
    End If
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = mList.IsFixedSize
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = mList.IsReadOnly
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Set mList(Index) = RHS
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    mList(Index) = RHS
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Call Helper.MoveVariant(ArrayList_Item, mList(Index))
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef startIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRangeReverse(mList.Count, startIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "Index", ElemCount, "Count", IsMissing(startIndex))
    
    If Comparer Is Nothing Then
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    Dim i As Long
    For i = ElemIndex To ElemIndex - ElemCount + 1 Step -1
        If Comparer.Compare(mList(i), Value) = 0 Then
            ArrayList_LastIndexOf = i
            Exit Function
        End If
    Next i
    
    ArrayList_LastIndexOf = -1
End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    Set ArrayList_NewEnum = CreateEnumerator(ArrayList_GetEnumerator)
End Function

Private Sub ArrayList_Remove(Value As Variant, Optional ByVal Comparer As IComparer)
    Call mList.Remove(Value, Comparer)
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Call mList.RemoveAt(Index)
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Dim Result As Long
    Result = VerifyListRange(mList.Count, Index, Count)
    If Result Then Call ThrowListRangeException(Result, Index, "Index", Count, "Count")
    
    Dim i As Long
    For i = 1 To Count
        Call mList.RemoveAt(Index)
    Next i
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mList.Count, Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))

    Dim t1 As Variant
    Dim t2 As Variant
    Dim i As Long
    Dim j As Long
    
    i = ElemIndex
    j = ElemIndex + ElemCount - 1
    
    Do While i < j
        Call Helper.MoveVariant(t1, mList(i))
        Call Helper.MoveVariant(t2, mList(j))
        
        If IsObject(t1) Then
            Set mList(j) = t1
        Else
            mList(j) = t1
        End If
        
        If IsObject(t2) Then
            Set mList(i) = t2
        Else
            mList(i) = t2
        End If
    
        i = i + 1
        j = j - 1
    Loop
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    If IsArray(c) Then
        Call SetRange(Index, c)
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw New ArgumentNullException
            
        If TypeOf c Is ICollection Then
            Call SetRange(Index, c)
        ElseIf TypeOf c Is IEnumerable Then
            Call SetRange(Index, c)
        ElseIf TypeOf c Is Collection Then
            Call SetRange(Index, c)
        Else
            Throw Cor.NewArgumentException("Only arrays, ICollection, Collection, and IEnumerable objects are supported.")
        End If
    Else
        Throw Cor.NewArgumentException("Only arrays, ICollection, Collection, and IEnumerable objects are supported.")
    End If
End Sub

Private Sub SetRange(ByVal Index As Long, ByRef c As Variant)
    Dim v As Variant
    For Each v In c
        If IsObject(v) Then
            Set mList(Index) = v
        Else
            mList(Index) = v
        End If
        
        Index = Index + 1
    Next v
End Sub

Private Sub ArrayList_Sort(Optional ByRef startIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim Result      As Long
    Result = GetOptionalListRange(mList.Count, startIndex, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "Index", ElemCount, "Count", IsMissing(startIndex))
        
    If Comparer Is Nothing Then
        Set mComparer = VBCorLib.Comparer.Default
    Else
        Set mComparer = Comparer
    End If
    
    Call QuickSortList(ElemIndex, ElemIndex + ElemCount - 1)
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As ciArrayTypes = 12&) As Variant
    Dim Ret As Variant
    Ret = cArray.CreateInstance(ArrayType, mList.Count)
    Call mList.CopyTo(Ret, 0)
    
    ArrayList_ToArray = Ret
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = IObject_ToString()
End Function

Private Sub ArrayList_TrimToSize()
    ' do nothing
End Sub

Private Property Get ArrayList_Version() As Long
    ' do nothing
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
    IObject_ToString = Object.ToString(mList, App)
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
    Call ArrayList_Clear
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


