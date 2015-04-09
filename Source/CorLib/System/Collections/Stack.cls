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
' Represents a first-in-last-out collection of values.
'
' @see Constructors
' @see ICollection
' @see IEnumerable
' @see ICloneable
'
Option Explicit
Implements IObject
Implements ICloneable
Implements IEnumerable
Implements ICollection

Private Const PROP_COUNT            As String = "Count"
Private Const PROP_CAPACITY         As String = "Capacity"
Private Const PROP_ITEMPREFIX       As String = "Item"

Private Const DEF_CAPACITY          As Long = 16

Private mItems()    As Variant
Private mCount      As Long
Private mVersion    As Long
Private mCapacity   As Long


''
' Removes all values from the collection.
'
' @remarks The count is reset to zero, but the capacity remains unchanged.
'
Public Sub Clear()
    ReDim mItems(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Returns a copy of this instance.
'
' @return The copy of this instance.
'
Public Function Clone() As Stack
    Set Clone = New Stack
    Clone.InitClone mItems, mCount
End Function

''
' Returns whether the collection contains a value.
'
' @param value The value to search for.
' @param comparer A user supplied comparer for the values.
' @return Boolean indicating if the value was found in the collection.
' @remarks The user can use a custom IComparer object for special comparison
' rules, such as finding user-defined types.
'
Public Function Contains(ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    Contains = (cArray.IndexOf(mItems, Value, 0, mCount, Comparer) >= 0)
End Function

''
' Copies the values to an array.
'
' @param dstArray The array to receive the values.
' @param index The starting position in dstArray to start placing the values.
'
Public Sub CopyTo(ByRef DstArray As Variant, ByVal Index As Long)
    Require.NotNullArray DstArray, Parameter_DstArray
    Require.That cArray.GetRank(DstArray) = 1, Rank_MultiDimNotSupported, Parameter_DstArray
    Require.Range Index >= LBound(DstArray), Parameter_Index, ArgumentOutOfRange_ArrayLB
    
    cArray.CopyEx mItems, 0, DstArray, Index, mCount
    cArray.Reverse DstArray, Index, mCount
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
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns an enumerator to enumerate the colleciton
'
' @return An enumerator.
'
Public Function GetEnumerator() As IEnumerator
    Dim Ret As New StackEnumerator
    Ret.Init Me, SAPtr(mItems)
    Set GetEnumerator = Ret
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
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
' Returns the next to be removed value from the stack without
' removing the value.
'
' @return The next value to be removed from the stack.
'
Public Function Peek() As Variant
    If mCount = 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EmptyStack))
    
    VariantCopy Peek, mItems(mCount - 1)
End Function

''
' Returns the next value from the stack, removing the item from the collection.
'
' @return The value just removed from the collection.
'
Public Function Pop() As Variant
    If mCount = 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EmptyStack))
    
    mCount = mCount - 1
    Helper.MoveVariant Pop, mItems(mCount)
    mVersion = mVersion + 1
End Function

''
' Places a value into the collection and becomes the next to be
' removed value from the collection.
'
' @param value The value to add to the collection
'
Public Sub Push(ByRef Value As Variant)
    EnsureCapacity mCount + 1
    VariantCopyInd mItems(mCount), Value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

''
' Returns a Variant array copy of the stack.
'
' @return A Variant type array containing all of the values in the collection.
' @remarks The order of the returned array is the first element corrisponds to
' the next value to be Popped off the stack.
'
Public Function ToArray() As Variant()
    Dim Result() As Variant
    Result = cArray.CreateInstance(vbVariant, mCount)
    
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
    ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub InitClone(ByRef OtherItems() As Variant, ByVal OtherCount As Long)
    mCount = OtherCount
    mItems = OtherItems
    mCapacity = UBound(mItems) + 1
End Sub

Friend Sub Init(ByRef Source As Variant)
    WriteRange Source
End Sub

Friend Property Get Version() As Long
    Version = mVersion
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WriteRange(ByRef Source As Variant)
    If IsArray(Source) Then
        Require.NotNullSingleDimensionArray Source
        WriteArray Source
    ElseIf IsObject(Source) Then
        Require.NotNothing Source, Parameter_Source, ArgumentNull_Collection
        
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
    FillStack Source, ArrayLength(Source)
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
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    EnsureCapacity DEF_CAPACITY
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        EnsureCapacity .ReadProperty(PROP_CAPACITY, DEF_CAPACITY)
        mCount = .ReadProperty(PROP_COUNT, 0)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Helper.MoveVariant mItems(i), .ReadProperty(PROP_ITEMPREFIX & i, Empty)
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_CAPACITY, mCapacity
        .WriteProperty PROP_COUNT, mCount
        
        Dim i As Long
        For i = 0 To mCount - 1
            .WriteProperty PROP_ITEMPREFIX & i, mItems(i)
        Next i
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function

