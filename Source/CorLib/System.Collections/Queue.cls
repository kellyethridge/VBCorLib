VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: Queue
'

''
' Represents a first-in, first-out collection of values.
'
' @include "..\DocIncludes\Queue.html"
Option Explicit
Implements IObject
Implements ICloneable
Implements ICollection
Implements IEnumerable
Implements IVersionable

Private Const DefaultCapacity       As Long = 16
Private Const ItemPropertyFormat    As String = "Item{0}"
Private Const HeadProperty          As String = "Head"
Private Const TailProperty          As String = "Tail"
Private Const CountProperty         As String = "Count"
Private Const CapacityProperty      As String = "Capacity"

Private mItems()    As Variant
Private mHead       As Long
Private mTail       As Long
Private mCapacity   As Long
Private mCount      As Long
Private mVersion    As Long


''
' Removes all items from the <a href="Queue.html">Queue</a>.
'
' @remarks <a href="Queue.Get.Count.html">Count</a> is set to zero, and references to other objects from elements of the collection are also released.
' @see #TrimToSize
' @include "..\DocIncludes\Queue.Clear.html"
Public Sub Clear()
    mCount = 0
    mHead = 0
    mTail = 0
    ReDim mItems(0 To mCapacity - 1)
    mVersion = mVersion + 1
End Sub

''
' Creates a shallow copy of the <a href="Queue.html">Queue</a>.
'
' @return A shallow copy of the <a href="Queue.html">Queue</a>.
' @remarks Returns a second queue with the same contents as the original.
' <p>A shallow copy of a collection copies only the elements of the collection, whether they are reference types or value types,
' but it does not copy the objects that the references refer to. The references in the new collection point to the same objects
' that the references in the original collection point to.</p>
Public Function Clone() As Queue
    Set Clone = New Queue
    Clone.InitClone mItems, mCount, mHead, mTail
End Function

''
' Determines whether an element is in the <a href="Queue.html">Queue</a>.
'
' @param Value The value to locate in the <a href="Queue.html">Queue</a>.
' @return <b>True</b> if <i>Value</i> is found in the <a href="Queue.html">Queue</a>; otherwise <b>False</b>.
' @remarks This method determines equality by calling <a href="ObjectStatic.Equals.html">Object.Equals</a>.
' <p>This method performs a linear search; therefore, this method is an O(n) operation, where n is <a href="Queue.Get.Count.html">Count</a>.</p>
Public Function Contains(ByRef Value As Variant) As Boolean
    Dim i As Long
    Dim j As Long
    
    j = mHead
    For i = 1 To mCount
        If Object.Equals(mItems(j), Value) Then
            Contains = True
            Exit Function
        End If
        j = (j + 1) Mod mCapacity
    Next i
End Function

''
' Copies the <a href="Queue.html">Queue</a> elements to an existing one-dimensional array, starting at the specified array index.
'
' @param Arr The one-dimensional array that is the destination of the elements copied from <a href="Queue.html">Queue</a>.
' @param Index The lowerbound-based index in <i>Arr</i> at which copying begins.
' @see #ToArray
' @include "..\DocIncludes\Queue.CopyTo.html"
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)
    ValidateArrayRange Arr, Index, mCount
        
    If mHead < mTail Then
        CorArray.CopyEx mItems, mHead, Arr, Index, mCount
    Else
        CorArray.CopyEx mItems, mHead, Arr, Index, mCapacity - mHead
        CorArray.CopyEx mItems, 0, Arr, (mCapacity - mHead) + Index, mTail
    End If
End Sub

''
' Gets the number of elements contained in the <a href="Queue.html">Queue</a>.
'
' @return The number of elements contained in the <a href="Queue.html">Queue</a>.
' @remarks The capacity of a <a href="Queue.html">Queue</a> is the number of elements that the <a href="Queue.html">Queue</a> can store.
' <b>Count</b> is the number of elements that are actually in the <a href="Queue.html">Queue</a>.
' <p>The capacity of a <a href="Queue.html">Queue</a> is always greater than or equal to <b>Count</b>. If <b>Count</b> exceeds the capacity while adding elements,
' the capacity is automatically increased by reallocating the internal array before copying the old elements and adding the new elements.</p>
' <p>The capacity can be decreased by calling <a href="Queue.TrimToSize.html">TrimToSize</a>.</p>
Public Property Get Count() As Long
    Count = mCount
End Property

''
' Removes and returns the item at the beginning of the <a href="Queue.html">Queue</a>.
'
' @return The object that is removed from the beginning of the <a href="Queue.html">Queue</a>.
' @see #Enqueue
' @see #Peek
' @include "..\DocIncludes\Queue.Dequeue.html"
Public Function Dequeue() As Variant
    If mCount = 0 Then _
        Error.InvalidOperation InvalidOperation_EmptyQueue
    
    Helper.MoveVariant Dequeue, mItems(mHead)
    mHead = (mHead + 1) Mod mCapacity
    mCount = mCount - 1
    mVersion = mVersion + 1
End Function

''
' Adds an item to the tail of the <a href="Queue.html">Queue</a>, expanding the queue as necessary.
'
' @param Value The item to be placed at the tail of the <a href="Queue.html">Queue</a>.
' @see #Dequeue
' @see #Peek
' @include "..\DocIncludes\Queue.Enqueue.html"
Public Sub Enqueue(ByRef Value As Variant)
    EnsureCapacity mCount + 1
    VariantCopyInd mItems(mTail), Value
    mTail = (mTail + 1) Mod mCapacity
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns an enumerator for the queue.
'
' @return Enumerator to allow iterating over the queue.
'
Public Function GetEnumerator() As IEnumerator
    Dim Ret As New QueueEnumerator
    Ret.Init Me, mHead, mItems
    Set GetEnumerator = Ret
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns an enumerator compatible with For..Each convention.
'
' @return The enumerator object.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Returns the item at the beginning of the <a href="Queue.html">Queue</a> without removing it.
'
' @return The item at the beginning of the <a href="Queue.html">Queue</a>.
' @see #Enqueue
' @see #Dequeue
' @include "..\DocIncludes\Queue.Peek.html"
Public Function Peek() As Variant
    If mCount = 0 Then _
        Error.InvalidOperation InvalidOperation_EmptyQueue
    
    VariantCopy Peek, mItems(mHead)
End Function

''
' Copies the <a href="Queue.html">Queue</a> elements to a new array.
'
' @return A new array containing elements copied from the <a href="Queue.html">Queue</a>.
' @see #CopyTo
' @include "..\DocIncludes\Queue.ToArray.html"
Public Function ToArray() As Variant()
    Dim Result() As Variant
    
    Result = CorArray.CreateInstance(vbVariant, mCount)
    If mCount > 0 Then
        CopyTo Result, 0
    End If
    
    ToArray = Result
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Sets the capacity to the actual number of elements in the <a href="Queue.html">Queue</a>.
'
' @remarks This method can be used to minimize a queue's memory overhead if no new elements will be added to the queue.
' <p>To reset a <a href="Queue.html">Queue</a> to its initial state, call the <a href="Queue.Clear.html">Clear</a> method before calling <b>TrimToSize</b>.
' Trimming an empty <a href="Queue.html">Queue</a> sets the capacity of the <a href="Queue.html">Queue</a> to the default capacity.</p>
' <p>This method is an O(<i>n</i>) operation, where <i>n</i> is <a href="Queue.Get.Count.html">Count</a>.</p>
' @see #Clear
' @see #Count
Public Sub TrimToSize()
    If mCount = 0 Then
        ReDim mItems(0 To DefaultCapacity - 1)
        mCapacity = DefaultCapacity
    Else
        NormalizeArray mCount
        mCapacity = mCount
    End If
    mVersion = mVersion + 1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    ReDim mItems(0 To DefaultCapacity - 1)
    mCapacity = DefaultCapacity
End Sub

Friend Sub InitClone(ByRef Items() As Variant, ByVal Count As Long, ByVal Head As Long, ByVal Tail As Long)
    mItems = Items
    mCount = Count
    mHead = Head
    mTail = Tail
    mCapacity = UBound(Items) + 1
End Sub

Friend Sub Init(ByRef Source As Variant)
    WriteRange Source
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NormalizeArray(ByVal NewCapacity As Long)
    Dim NewItems() As Variant
    
    If mHead = 0 Then
        If NewCapacity <> mCapacity Then
            ReDim Preserve mItems(0 To NewCapacity - 1)
        End If
    Else
        ReDim NewItems(0 To NewCapacity - 1)
        If mHead < mTail Then
            CopyMemory NewItems(0), mItems(mHead), (mTail - mHead) * vbSizeOfVariant
        Else
            CopyMemory NewItems(0), mItems(mHead), (mCapacity - mHead) * vbSizeOfVariant
            CopyMemory NewItems(mCapacity - mHead), mItems(0), mHead * vbSizeOfVariant
        End If
        
        ZeroMemory mItems(0), mCapacity * vbSizeOfVariant
        Helper.Swap4 ByVal ArrPtr(mItems), ByVal ArrPtr(NewItems)
    End If
    mTail = mCount
    mHead = 0
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    
    If RequiredCapacity <= mCapacity Then
        Exit Sub
    End If
    NewCapacity = mCapacity * 2
    
    If RequiredCapacity > NewCapacity Then
        NewCapacity = RequiredCapacity
    End If
    
    NormalizeArray NewCapacity
    mCapacity = NewCapacity
End Sub

Private Sub WriteRange(ByRef Source As Variant)
    If IsArray(Source) Then
        WriteArray Source
    ElseIf IsObject(Source) Then
        If Source Is Nothing Then _
            Error.ArgumentNull "Source", ArgumentNull_Collection
        
        If TypeOf Source Is Collection Then
            WriteVBCollection Source
        ElseIf TypeOf Source Is ICollection Then
            WriteICollection Source
        Else
            Throw Cor.NewInvalidCastException("An ICollection or VBA.Collection object is required.")
        End If
    Else
        Throw Cor.NewInvalidCastException("An Array, ICollection, or VBA.Collection object is required.")
    End If
End Sub

Private Sub WriteArray(ByRef Source As Variant)
    If CorArray.IsNull(Source) Then _
        Error.ArgumentNull "Source", ArgumentNull_Array
    
    WriteCollection Len1D(Source), Source
End Sub

Private Sub WriteVBCollection(ByVal Source As Collection)
    WriteCollection Source.Count, Source
End Sub

Private Sub WriteICollection(ByVal Source As ICollection)
    WriteCollection Source.Count, Source
End Sub

Private Sub WriteCollection(ByVal SizeOfCollection As Long, ByRef Source As Variant)
    EnsureCapacity SizeOfCollection
        
    Dim Value As Variant
    Dim Index As Long
    For Each Value In Source
        Helper.MoveVariant mItems(Index), Value
        Index = Index + 1
    Next
    
    mTail = SizeOfCollection
    mCount = SizeOfCollection
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mCount = .ReadProperty(CountProperty, 0)
        mHead = .ReadProperty(HeadProperty, 0)
        mTail = .ReadProperty(TailProperty, 0)
        EnsureCapacity .ReadProperty(CapacityProperty, DefaultCapacity)
        ReDim mItems(0 To mCapacity - 1)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Helper.MoveVariant mItems(i), .ReadProperty(MakeItemProperty(i), Empty)
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty CountProperty, mCount
        .WriteProperty HeadProperty, mHead
        .WriteProperty TailProperty, mTail
        .WriteProperty CapacityProperty, mCapacity
        
        NormalizeArray mCapacity
        
        Dim i As Long
        For i = 0 To mCount - 1
            .WriteProperty MakeItemProperty(i), mItems(i)
        Next i
    End With
End Sub

Private Function MakeItemProperty(ByVal Number As Long) As String
    MakeItemProperty = CorString.Format(ItemPropertyFormat, Number)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long)
    Call CopyTo(Arr, Index)
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = GetEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IVersionableCollection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IVersionable_Version() As Long
    IVersionable_Version = mVersion
End Property
