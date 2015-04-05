VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IListWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: IListWrapper
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
Friend Sub Init(ByVal List As IList)
    Set mList = List
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub QuickSortList(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim x As Variant
    
    Do While Left < Right
        i = Left
        j = Right
        Helper.MoveVariant x, mList((i + j) \ 2)
        
        Do
            Do While mComparer.Compare(mList(i), x) < 0
                i = i + 1
            Loop
            Do While mComparer.Compare(mList(j), x) > 0
                j = j - 1
            Loop
            
            If i > j Then Exit Do
            
            If i < j Then
                Dim t1 As Variant
                Dim t2 As Variant
                
                Helper.MoveVariant t1, mList(i)
                Helper.MoveVariant t2, mList(j)
                
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
                
            i = i + 1
            j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then
                QuickSortList Left, j
            End If
            Left = i
        Else
            If i < Right Then
                QuickSortList i, Right
            End If
            Right = j
        End If
    Loop
End Sub

Private Sub InsertItems(ByVal Index As Long, Items As Variant)
    Dim Item As Variant
    
    For Each Item In Items
        mList.Insert Index, Item
        Index = Index + 1
    Next
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    ArrayList_Add = mList.Add(Value)
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Dim Value As Variant
    
    For Each Value In c
        mList.Add Value
    Next
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_BinarySearch = ArrayList_BinarySearchEx(0, mList.Count, Value, Comparer)
End Function

Private Function ArrayList_BinarySearchEx(ByVal Index As Long, ByVal Count As Long, Value As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim MiddleIndex As Long
    Dim LowIndex    As Long
    Dim HighIndex   As Long
    
    CheckRange Index < 0, Param.Index, ArgumentOutOfRange.NeedNonNegNum
    CheckRange Count < 0, Param.Count, ArgumentOutOfRange.NeedNonNegNum
    CheckArgument Index + Count > mList.Count, Argument_InvalidOffLen
    
    If Comparer Is Nothing Then
        Set Comparer = VBCorLib.Comparer.Default
    End If
    
    LowIndex = Index
    HighIndex = LowIndex + Count - 1
    Do While LowIndex <= HighIndex
        MiddleIndex = (LowIndex + HighIndex) \ 2
        Select Case Comparer.Compare(mList(MiddleIndex), Value)
            Case 0
                ArrayList_BinarySearchEx = MiddleIndex
                Exit Function
            Case Is > 0
                HighIndex = MiddleIndex - 1
            Case Else
                LowIndex = MiddleIndex + 1
        End Select
    Loop
    
    ArrayList_BinarySearchEx = Not LowIndex
End Function

Private Property Let ArrayList_Capacity(ByVal Value As Long)
    If Value < mList.Count Then
        Throw Cor.NewArgumentOutOfRangeException("Value", Value, Resources.GetString(ArgumentOutOfRange_SmallCapacity))
    End If
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.Count
End Property

Private Sub ArrayList_Clear()
    mList.Clear
End Sub

Private Function ArrayList_Clone() As ArrayList
    Dim Wrapper As New IListWrapper
    Wrapper.Init mList
    Set ArrayList_Clone = Wrapper
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    ArrayList_Contains = mList.Contains(Value, Comparer)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional ArrayIndex As Variant)
    Dim Index As Long
    Index = GetOptionalLong(ArrayIndex, LBound(DstArray))
    
    If Index < LBound(DstArray) Then _
        Throw Cor.NewArgumentOutOfRangeException("ArrayIndex", ArrayIndex, Resources.GetString(ArgumentOutOfRange_LBound))
    
    mList.CopyTo DstArray, Index
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mList.Count Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
    
    Dim Result As Long
    Result = VerifyArrayRange(GetArrayPointer(DstArray), ArrayIndex, Count)
    If Result Then _
        ThrowArrayRangeException Result, "DstArray", ArrayIndex, "ArrayIndex", Count, "Count"
    
    Dim v As Variant
    Dim i As Long
    
    For i = 0 To Count - 1
        Helper.MoveVariant v, mList(Index + i)
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

Private Function ArrayList_GetEnumerator(Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Object
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    ElemIndex = GetOptionalLong(Index, 0)
    ElemCount = GetOptionalLong(Count, mList.Count - ElemIndex)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", ElemIndex, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mList.Count Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
    
    Dim Ret As New ArrayListEnumerator
    Ret.Init Me, ElemIndex, ElemCount
    Set ArrayList_GetEnumerator = Ret
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = IObject_GetHashcode
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Count, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mList.Count Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
    
    Dim Ret As RangedArrayList
    Set Ret = New RangedArrayList
    Ret.Init Me, Index, Count
    Set ArrayList_GetRange = Ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    ElemIndex = GetOptionalLong(StartIndex, 0)
    ElemCount = GetOptionalLong(Count, mList.Count - ElemIndex)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", ElemIndex, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex >= mList.Count Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", ElemIndex, Resources.GetString(ArgumentOutOfRange_LargerThanCollection))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mList.Count Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Resources.GetString(ArgumentOutOfRange_LargerThanCollection))
    
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
    mList.Insert Index, Value
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    If IsArray(c) Then
        InsertItems Index, c
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw New ArgumentNullException
            
        If TypeOf c Is ICollection Then
            InsertItems Index, c
        ElseIf TypeOf c Is IEnumerable Then
            InsertItems Index, c
        ElseIf TypeOf c Is Collection Then
            InsertItems Index, c
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
    Helper.MoveVariant ArrayList_Item, mList(Index)
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    ElemIndex = GetOptionalLong(StartIndex, mList.Count - 1)
    ElemCount = GetOptionalLong(Count, ElemIndex + 1)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", Message:=Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex >= mList.Count Then _
        Throw Cor.NewArgumentOutOfRangeException("StartIndex", Message:=Resources.GetString(ArgumentOutOfRange_LargerThanCollection))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex - ElemCount + 1 < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Resources.GetString(ArgumentOutOfRange_LargerThanCollection))
    
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
    mList.Remove Value, Comparer
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    mList.RemoveAt Index
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Count, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If Index + Count > mList.Count Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
    
    Dim i As Long
    For i = 1 To Count
        mList.RemoveAt Index
    Next i
End Sub

Private Sub ArrayList_Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    If GetOptionalLongPair(Index, 0, ElemIndex, Count, mList.Count, ElemCount) = Argument_ParamRequired Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_ParamRequired), IIf(IsMissing(Index), "Index", "Count"))
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", ElemIndex, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mList.Count Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))

    Dim t1 As Variant
    Dim t2 As Variant
    Dim i As Long
    Dim j As Long
    
    i = ElemIndex
    j = ElemIndex + ElemCount - 1
    
    Do While i < j
        Helper.MoveVariant t1, mList(i)
        Helper.MoveVariant t2, mList(j)
        
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
        SetRange Index, c
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw New ArgumentNullException
            
        If TypeOf c Is ICollection Then
            SetRange Index, c
        ElseIf TypeOf c Is IEnumerable Then
            SetRange Index, c
        ElseIf TypeOf c Is Collection Then
            SetRange Index, c
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

Private Sub ArrayList_Sort(Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    
    ElemIndex = GetOptionalLong(Index, 0)
    ElemCount = GetOptionalLong(Count, mList.Count - ElemIndex)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", ElemIndex, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", ElemCount, Resources.GetString(ArgumentOutOfRange_NeedNonNegNum))
    If ElemIndex + ElemCount > mList.Count Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_InvalidOffLen))
        
    If Comparer Is Nothing Then
        Set mComparer = VBCorLib.Comparer.Default
    Else
        Set mComparer = Comparer
    End If
    
    QuickSortList ElemIndex, ElemIndex + ElemCount - 1
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As vbVarType = 12&) As Variant
    ArrayList_ToArray = cArray.CreateInstance(ArrayType, mList.Count)
    mList.CopyTo ArrayList_ToArray, 0
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

Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    IList_Contains = ArrayList_Contains(Value, Comparer)
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

Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long
    IList_IndexOf = ArrayList_IndexOf(Value, , , Comparer)
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

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant)
    Set ArrayList_Item(Index) = RHS
End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant)
    ArrayList_Item(Index) = RHS
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant IList_Item, ArrayList_Item(Index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = ArrayList_NewEnum
End Function

Private Sub IList_Remove(Value As Variant, Optional ByVal Comparer As IComparer)
    ArrayList_Remove Value, Comparer
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    ArrayList_RemoveAt Index
End Sub


