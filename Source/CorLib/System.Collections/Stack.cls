VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: Stack
'

''
' Represents a simple last-in-first-out (LIFO) non-generic collection of items.
'
' @see Constructors
' @see ICollection
' @see IEnumerable
' @see ICloneable
' @include "..\DocIncludes\Stack.html"
Option Explicit
Implements IObject
Implements ICloneable
Implements IEnumerable
Implements ICollection
Implements IVersionable

Private Const CountProperty         As String = "Count"
Private Const CapacityProperty      As String = "Capacity"
Private Const ItemPropertyFormat    As String = "Item{0}"
Private Const DefaultCapacity       As Long = 16

Private mItems()    As Variant
Private mCount      As Long
Private mVersion    As Long
Private mCapacity   As Long


''
' Removes all items from the <a href="Stack.html">Stack</a>.
'
' @include "..\DocIncludes\Stack.Clear.html"
Public Sub Clear()
    ReDim mItems(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Creates a shallow copy of the <a href="Stack.html">Stack</a>.
'
' @return A shallow copy of the <a href="Stack.html">Stack</a>.
' @remarks Returns a second stack with the same contents as the original.
' <p>A shallow copy of a collection copies only the elements of the collection, whether they are reference types or value types,
' but it does not copy the objects that the references refer to. The references in the new collection point to the same objects
' that the references in the original collection point to.</p>
Public Function Clone() As Stack
    Set Clone = New Stack
    Clone.InitClone mItems, mCount
End Function

''
' Determines whether an element is in the <a href="Stack.html">Stack</a>.
'
' @param Value The value to locate in the <a href="Stack.html">Stack</a>.
' @return <b>True</b> if <i>Value</i> is found in the <a href="Stack.html">Stack</a>; otherwise, <b>False</b>.
' @remarks This method determines equality by calling <a href="ObjectStatic.Equals.html">Object.Equals</a>.
' <p>This method performs a linear search; therefore, this method is an O(<i>n</i>) operation, where <i>n</i> is <a href="Stack.Get.Count.html">Count</a>.
Public Function Contains(ByRef Value As Variant) As Boolean
    Contains = CorArray.IndexOf(mItems, Value, 0, mCount) >= 0
End Function

''
' Copies the <a href="Stack.html">Stack</a> to an existing one-dimensional array, starting at the specified array index.
'
' @param Arr The one-dimensional array that is the destination of the elements copied from <a href="Stack.html">Stack</a>.
' @param index The lowerbound-based index in array at which copying begins.
' @see #ToArray
' @include "..\DocIncludes\Stack.CopyTo.html"
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)
    If CorArray.IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    If CorArray.Rank(Arr) <> 1 Then _
        Error.Argument Argument_MultiDimNotSupported, "Arr"  ' .NET throws an ArgumentException instead of RankException.
    
    If Index < LBound(Arr) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayLB
    
    CorArray.CopyEx mItems, 0, Arr, Index, mCount
    CorArray.Reverse Arr, Index, mCount
End Sub

''
' Returns the number of items on the stack.
'
' @return The number of items.
'
Public Property Get Count() As Long
    Count = mCount
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare this object to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns an enumerator to enumerate the colleciton
'
' @return An enumerator.
'
Public Function GetEnumerator() As IEnumerator
    Dim Ret As New StackEnumerator
    Ret.Init Me, mItems
    Set GetEnumerator = Ret
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns an enumerator to be used in For..Each loops
'
' @return An enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Returns the item at the top of the <a href="Stack.html">Stack</a> without removing it.
'
' @return The item at the top of the <a href="Stack.html">Stack</a>.
' @see #Pop
' @see #Push
' @include "..\DocIncludes\Stack.Peek.html"
Public Function Peek() As Variant
    If mCount = 0 Then _
        Error.InvalidOperation InvalidOperation_EmptyStack
    
    VariantCopy Peek, mItems(mCount - 1)
End Function

''
' Removes and returns the item at the top of the <a href="Stack.html">Stack</a>.
'
' @return The item removed from the top of the <a href="Stack.html">Stack</a>.
' @see #Push
' @see #Peek
' @include "..\DocIncludes\Stack.Pop.html"
Public Function Pop() As Variant
    If mCount = 0 Then _
        Error.InvalidOperation InvalidOperation_EmptyStack
    
    mCount = mCount - 1
    Helper.MoveVariant Pop, mItems(mCount)
    mVersion = mVersion + 1
End Function

''
' Inserts an item at the top of the <a href="Stack.html">Stack</a>.
'
' @param value The item to push onto the <a href="Stack.html">Stack</a>.
' @see #Pop
' @see #Peek
' @include "..\DocIncludes\Stack.Push.html"
Public Sub Push(ByRef Value As Variant)
    EnsureCapacity mCount + 1
    VariantCopyInd mItems(mCount), Value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

''
' Copies the <a href="Stack.html">Stack</a> to a new array.
'
' @return A new array containing copies of the elements of the <a href="Stack.html">Stack</a>.
' @see #CopyTo
' @include "..\DocIncludes\Stack.ToArray.html"
Public Function ToArray() As Variant()
    Dim Result() As Variant
    Result = CorArray.CreateInstance(vbVariant, mCount)
    
    Dim i As Long
    For i = 0 To mCount - 1
        VariantCopy Result(i), mItems(mCount - i - 1)
    Next i
    
    ToArray = Result
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    EnsureCapacity DefaultCapacity
End Sub

Friend Sub InitClone(ByRef OtherItems() As Variant, ByVal OtherCount As Long)
    mCount = OtherCount
    mItems = OtherItems
    mCapacity = UBound(mItems) + 1
End Sub

Friend Sub Init(ByRef Source As Variant)
    WriteRange Source
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WriteRange(ByRef Source As Variant)
    If IsArray(Source) Then
        If CorArray.IsNull(Source) Then _
            Error.ArgumentNull "Source", ArgumentNull_Array
        If CorArray.Rank(Source) <> 1 Then _
            Error.Argument Argument_MultiDimNotSupported, "Source"
        
        WriteArray Source
    ElseIf IsObject(Source) Then
        If Source Is Nothing Then _
            Error.ArgumentNull "Source", ArgumentNull_Collection
        
        If TypeOf Source Is Collection Then
            WriteCollection Source
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
    FillStack Source, Len1D(Source)
End Sub

Private Sub WriteCollection(ByVal Source As Collection)
    FillStack Source, Source.Count
End Sub

Private Sub WriteICollection(ByVal Source As ICollection)
    FillStack Source, Source.Count
End Sub

Private Sub FillStack(ByRef Source As Variant, ByVal SourceSize As Long)
    EnsureCapacity SourceSize
        
    Dim Value As Variant
    Dim Index As Long
    For Each Value In Source
        Helper.MoveVariant mItems(Index), Value
        Index = Index + 1
    Next
    
    mCount = SourceSize
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    If RequiredCapacity <= mCapacity Then
        Exit Sub
    End If
    
    Dim NewCapacity As Long
    NewCapacity = mCapacity * 2
    
    If RequiredCapacity > NewCapacity Then
        NewCapacity = RequiredCapacity
    End If
    
    ReDim Preserve mItems(0 To NewCapacity - 1)
    mCapacity = NewCapacity
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        EnsureCapacity .ReadProperty(CapacityProperty, DefaultCapacity)
        mCount = .ReadProperty(CountProperty, 0)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Helper.MoveVariant mItems(i), .ReadProperty(MakeItemProperty(i), Empty)
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty CapacityProperty, mCapacity
        .WriteProperty CountProperty, mCount
        
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
    CopyTo Arr, Index
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
