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
'    Module: Queue
'

''
' Provides a collections that maintains a list of items in a First-in-First-Out
' sequence.
'
' @see Constructors
' @see ICollection
' @see IEnumerable
' @see ICloneable
'
Option Explicit
Implements IObject
Implements ICloneable
Implements ICollection
Implements IEnumerable
Implements IVersionableCollection

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
' Clears the contents of the queue
'
' @remarks Even though the contents are cleared, the capacity
' of the queue is unchanged.
'
Public Sub Clear()
    mCount = 0
    mHead = 0
    mTail = 0
    ReDim mItems(0 To mCapacity - 1)
    mVersion = mVersion + 1
End Sub

''
' Returns a clone of the queue.
'
' @remarks Returns a second queue with the same contents as the original.
'
Public Function Clone() As Queue
    Set Clone = New Queue
    Clone.InitClone mItems, mCount, mHead, mTail
End Function

''
' Determines if the queue contains a specific value.
'
' @param value The value to search for in the queue.
' @return Value indicating if the value was found.
' @remarks If more than one instance of the same value exists in the
' queue, this will stop once it finds the first instance.
'
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
' Copy the contents of the queue to an array.
'
' @param Arr The array to copy the contents to.
' @param Index The starting index in the dstArray to copy the contents to.
' @remarks The values in the queue will be cast to the array datatype if possible.
'
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
' Returns the number of items in the queue.
'
' @return The number of items in the queue.
'
Public Property Get Count() As Long
    Count = mCount
End Property

''
' Returns the item at the head of the queue and removes that item from the queue.
'
' @return The item at the head of the queue.
'
Public Function Dequeue() As Variant
    If mCount = 0 Then _
        Error.InvalidOperation InvalidOperation_EmptyQueue
    
    Helper.MoveVariant Dequeue, mItems(mHead)
    mHead = (mHead + 1) Mod mCapacity
    mCount = mCount - 1
    mVersion = mVersion + 1
End Function

''
' Adds an item to the tail of the queue, expanding the queue as necessary.
'
' @param value The item to be placed at the tail of the queue.
'
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
' Returns the next value to be removed from the queue without
' removing the value from the queue.
'
' @return The next value in the queue.
'
Public Function Peek() As Variant
    If mCount = 0 Then _
        Error.InvalidOperation InvalidOperation_EmptyQueue
    
    VariantCopy Peek, mItems(mHead)
End Function

''
' Returns the contents of the queue in a variant array.
'
' @return An array of variants containing the queue contents.
'
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
' Resizes the internal capacity to match the number of items in the queue.
'
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
    
    WriteCollection CorArray.LengthFirstDim(Source), Source
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
Private Property Get IVersionableCollection_CurrentVersion() As Long
    IVersionableCollection_CurrentVersion = mVersion
End Property
