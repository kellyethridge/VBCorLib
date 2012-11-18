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
'    Module: Stack
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
    Call Clone.CloneHelper(mItems, mCount)
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
    Call cArray.CopyEx(mItems, 0, DstArray, Index, mCount)
    Call cArray.Reverse(DstArray, Index, mCount)
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
    Call Ret.Init(Me, SAPtr(mItems))
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
    
    Call VariantCopy(Peek, mItems(mCount - 1))
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
    Call Helper.MoveVariant(Pop, mItems(mCount))
    mVersion = mVersion + 1
End Function

''
' Places a value into the collection and becomes the next to be
' removed value from the collection.
'
' @param value The value to add to the collection
'
Public Sub Push(ByRef Value As Variant)
    Call EnsureCapacity(mCount + 1)
    Call VariantCopyInd(mItems(mCount), Value)
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
    Dim Ret() As Variant
    Ret = cArray.CreateInstance(vbVariant, mCount)
    
    Dim i As Long
    For i = 0 To mCount - 1
        Call VariantCopy(Ret(i), mItems(mCount - i - 1))
    Next i
    ToArray = Ret
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
Friend Sub CloneHelper(ByRef Items() As Variant, ByVal Count As Long)
    mCount = Count
    mItems = Items
    mCapacity = UBound(Items) + 1
End Sub

Friend Sub Init(ByRef c As Variant)
    Call WriteRange(c)
End Sub

Friend Property Get Version() As Long
    Version = mVersion
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WriteRange(ByRef c As Variant)
    If IsArray(c) Then
        Call WriteArray(c)
    ElseIf IsObject(c) Then
        If c Is Nothing Then _
            Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Collection), "c")
        
        If TypeOf c Is Collection Then
            Call WriteVBCollection(c)
        ElseIf TypeOf c Is ICollection Then
            Call WriteICollection(c)
        Else
            Throw Cor.NewInvalidCastException("An ICollection or VBA.Collection object is required.")
        End If
    Else
        Throw Cor.NewInvalidCastException("An Array, ICollection, or VBA.Collection object is required.")
    End If

End Sub

Private Sub WriteArray(ByRef Arr As Variant)
    If cArray.IsNull(Arr) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "c")
    
    Call WriteCollection(cArray.GetLength(Arr), Arr)
End Sub

Private Sub WriteVBCollection(ByVal CollectionToInsert As Collection)
    Call WriteCollection(CollectionToInsert.Count, CollectionToInsert)
End Sub

Private Sub WriteICollection(ByVal CollectionToInsert As ICollection)
    Call WriteCollection(CollectionToInsert.Count, CollectionToInsert)
End Sub

Private Sub WriteCollection(ByVal SizeOfCollection As Long, ByRef CollectionToWrite As Variant)
    Call EnsureCapacity(SizeOfCollection)
        
    Dim Value As Variant
    Dim Index As Long
    For Each Value In CollectionToWrite
        Call Helper.MoveVariant(mItems(Index), Value)
        Index = Index + 1
    Next Value
    mCount = SizeOfCollection
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    
    If RequiredCapacity <= mCapacity Then Exit Sub
    NewCapacity = mCapacity * 2
    If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
    ReDim Preserve mItems(0 To NewCapacity - 1)
    mCapacity = NewCapacity
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    ReDim mItems(0 To DEF_CAPACITY - 1)
    mCapacity = DEF_CAPACITY
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Call EnsureCapacity(.ReadProperty(PROP_CAPACITY, DEF_CAPACITY))
        mCount = .ReadProperty(PROP_COUNT, 0)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Call Helper.MoveVariant(mItems(i), .ReadProperty(PROP_ITEMPREFIX & i, Empty))
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_CAPACITY, mCapacity)
        Call .WriteProperty(PROP_COUNT, mCount)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Call .WriteProperty(PROP_ITEMPREFIX & i, mItems(i))
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function

