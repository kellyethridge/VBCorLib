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
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: IListWrapper
'

Option Explicit
Implements ArrayList

Private mList       As IList
Private mVersion    As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As IList)
    Set mList = List
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArrayList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ArrayList_Add(Value As Variant) As Long
    ArrayList_Add = mList.Add(Value)
    mVersion = mVersion + 1
End Function

Private Sub ArrayList_AddRange(c As Variant)
    Dim v As Variant
    For Each v In c
        Call mList.Add(v)
    Next v
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_BinarySearch(Value As Variant, Optional Index As Variant, Optional Count As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim HighIndex   As Long
    Dim LowIndex    As Long
    Dim MiddleIndex As Long
    
    If Comparer Is Nothing Then Set Comparer = modStaticClasses.Comparer.Default
    
    HighIndex = mList.Count - 1
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
    If RHS < mList.Count Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_SmallCapacity))
        
End Property

Private Property Get ArrayList_Capacity() As Long
    ArrayList_Capacity = mList.Count
End Property

Private Sub ArrayList_Clear()
    Call mList.Clear
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_Clone() As ArrayList
    Dim Ret As New IListWrapper
    Call Ret.Init(mList)
    Set ArrayList_Clone = Ret
End Function

Private Function ArrayList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    ArrayList_Contains = mList.Contains(Value, Comparer)
End Function

Private Sub ArrayList_CopyTo(DstArray As Variant, Optional ArrayIndex As Variant)
    Call mList.CopyTo(DstArray, ArrayIndex)
End Sub

Private Sub ArrayList_CopyToEx(ByVal Index As Long, DstArray As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    Dim DstArrPtr As Long
    DstArrPtr = GetArrayPointer(DstArray, True)
    
    Dim Result As Long
    Result = VerifyArrayRange(DstArrPtr, ArrayIndex, Count)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "DstArray", ArrayIndex, "ArrayIndex", Count, "Count", False)
    
    If Index < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Index")
    If Index + Count > mList.Count Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidCountOffset), "Index")
    
    Dim i As Long
    Dim v As Variant
    For i = Index To Index + Count - 1
        Call Helper.MoveVariant(v, mList(i))
        If IsObject(v) Then
            Set DstArray(i + ArrayIndex) = v
        Else
            DstArray(i + ArrayIndex) = v
        End If
    Next i
    
End Sub

Private Property Get ArrayList_Count() As Long
    ArrayList_Count = mList.Count
End Property

Private Function ArrayList_Equals(Value As Variant) As Boolean
    ArrayList_Equals = Object.Equals(Me, Value)
End Function

Private Function ArrayList_GetEnumerator(Optional StartIndex As Variant, Optional Count As Variant) As Object
    Set ArrayList_GetEnumerator = mList.GetEnumerator
End Function

Private Function ArrayList_GetHashCode() As Long
    ArrayList_GetHashCode = ObjPtr(CUnk(Me))
End Function

Private Function ArrayList_GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    Dim Result As Long
    Result = VerifyListRange(mList.Count, Index, Count)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, Index, "Index", Count, "Count")
    
    Dim Ret As New RangedArrayList
    Call Ret.Init(Me, Index, Count)
    Set ArrayList_GetRange = Ret
End Function

Private Function ArrayList_IndexOf(Value As Variant, Optional StartIndex As Variant, Optional Count As Variant, Optional ByVal Comparer As IComparer) As Long
    ArrayList_IndexOf = mList.IndexOf(Value, Comparer)
End Function

Private Sub ArrayList_Insert(ByVal Index As Long, Value As Variant)
    Call mList.Insert(Index, Value)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_InsertRange(ByVal Index As Long, c As Variant)
    Dim v As Variant
    For Each v In c
        Call mList.Insert(Index, v)
        Index = Index + 1
    Next v
    mVersion = mVersion + 1
End Sub

Private Property Get ArrayList_IsFixedSize() As Boolean
    ArrayList_IsFixedSize = mList.IsFixedSize
End Property

Private Property Get ArrayList_IsReadOnly() As Boolean
    ArrayList_IsReadOnly = mList.IsReadOnly
End Property

Private Property Set ArrayList_Item(ByVal Index As Long, RHS As Variant)
    Set mList(Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Let ArrayList_Item(ByVal Index As Long, RHS As Variant)
    mList(Index) = RHS
    mVersion = mVersion + 1
End Property

Private Property Get ArrayList_Item(ByVal Index As Long) As Variant
    Call Helper.MoveVariant(ArrayList_Item, mList(Index))
End Property

Private Function ArrayList_LastIndexOf(Value As Variant, Optional StartIndex As Variant, Optional Count As Variant, Optional ByVal Comparer As IComparer) As Long

End Function

Private Function ArrayList_NewEnum() As stdole.IUnknown
    
End Function

Private Sub ArrayList_Remove(Value As Variant, Optional ByVal Comparer As IComparer)
    Call mList.Remove(Value, Comparer)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_RemoveAt(ByVal Index As Long)
    Call mList.RemoveAt(Index)
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Do While Count > 0
        Call mList.RemoveAt(Index)
        Count = Count - 1
    Loop
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Reverse(Optional Index As Variant, Optional Count As Variant)
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    
    Result = GetOptionalLongPair(Index, 0, ElemIndex, Count, mList.Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowListRangeException(Result, ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim i As Long
    Dim j As Long
    Dim t1 As Variant
    Dim t2 As Variant
    
    i = ElemIndex
    j = ElemIndex + ElemCount - 1
    
    Do While i < j
        Call Helper.MoveVariant(t1, mList(i))
        Call Helper.MoveVariant(t2, mList(j))
        
        If IsObject(t2) Then
            Set mList(i) = t2
        Else
            mList(i) = t2
        End If
        
        If IsObject(t1) Then
            Set mList(j) = t1
        Else
            mList(j) = t1
        End If
        
        i = i + 1
        j = j - 1
    Loop
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_SetRange(ByVal Index As Long, c As Variant)
    Dim Size As Long
    
    If IsArray(c) Then
        If cArray.IsNull(c) Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "c")
        
            Size = cArray.GetLength(c)
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Collection), "c")
            
        If TypeOf c Is ICollection Then
            Dim ICol As ICollection
            Set ICol = c
            Size = ICol.Count
        ElseIf TypeOf c Is Collection Then
            Dim Col As Collection
            Set Col = c
            Size = Col.Count
        Else
            Throw Cor.NewArgumentException("Must be an array, collection or ICollection object.", "c")
        End If
    End If
    
    If (Index < 0) Or (Index + Size > mList.Count) Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "Index", Index)
        
    Dim v As Variant
    For Each v In c
        If IsObject(v) Then
            Set mList(Index) = v
        Else
            mList(Index) = v
        End If
        
        Index = Index + 1
    Next v
    mVersion = mVersion + 1
End Sub

Private Sub ArrayList_Sort(Optional StartIndex As Variant, Optional Count As Variant, Optional ByVal Comparer As IComparer)
    If mList.Count = 0 Then Exit Sub
    
    Dim Values() As Variant
    ReDim Values(0 To mList.Count - 1)
    Call mList.CopyTo(Values, 0)
    Call cArray.Sort(Values, Comparer)
    
    Dim i As Long
    For i = 0 To UBound(Values)
        If IsObject(Values(i)) Then
            Set mList(i) = Values(i)
        Else
            mList(i) = Values(i)
        End If
    Next i
    mVersion = mVersion + 1
End Sub

Private Function ArrayList_ToArray(Optional ByVal ArrayType As ciArrayTypes = 12&) As Variant
    ArrayList_ToArray = cArray.CreateInstance(ArrayType, mList.Count)
    Call mList.CopyTo(ArrayList_ToArray, 0)
End Function

Private Function ArrayList_ToString() As String
    ArrayList_ToString = Object.ToString(Me, App)
End Function

Private Sub ArrayList_TrimToSize()
    ' do nothing
End Sub

Private Property Get ArrayList_Version() As Long
    ArrayList_Version = mVersion
End Property
