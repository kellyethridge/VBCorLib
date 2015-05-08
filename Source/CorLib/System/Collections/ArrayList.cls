VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: ArrayList
'

''
' A list that dynamically increases in size to hold a number of items.
'
' @remarks This class contains an internal array of Variants. As new items are
' added to the list, the capacity will increase as necessary.
' <p>The capacity of the list is the current number of elements in the internal array.</p>
' <p>To lower the capacity to save memory, use the TrimToSize method, or set Capacity directly.</p>
' <p>The default capacity is 16. The Item index is zero based.</p>
' <p>This class is persistable.</p>
'
' @see Constructors
' @see ArrayListStatic
' @see IList
'
Option Explicit
Implements IObject
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IList
Implements IVersionableCollection

Private Const PROP_COUNT            As String = "Count"
Private Const PROP_CAPACITY         As String = "Capacity"
Private Const PROP_ITEMPREFIX       As String = "Item"
Private Const PROP_SUBTYPE          As String = "SubType"
Private Const PROP_DEFAULTCOMPARER  As String = "DefaultComparer"
Private Const PROP_COMPARER         As String = "Comparer"

Private Const SUBTYPE_NORMAL        As Long = 0
Private Const SUBTYPE_EMPTY         As Long = 1
Private Const SUBTYPE_NOTHING       As Long = 2
Private Const DEF_CAPACITY          As Long = 16

Private mItems()    As Variant
Private mCount      As Long
Private mCapacity   As Long
Private mVersion    As Long
Private mComparer   As IComparer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   API
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' Adds a new item to the end of the list.
'
' @param value The item to be added to the list.
' @return The index at which the item was added.
' @remarks As items are added, the capacity is increased as necessary. The items
' are appended to the end of the list and maintain the order they were added,
' provided no other method is used to change the order, such as Insert or Remove.
' @include "..\..\Includes\ArrayListAdd.txt"
Public Function Add(ByRef Value As Variant) As Long
    Add = mCount
    Insert mCount, Value
End Function

''
' Adds the items from a collection to the list.
'
' @param Source The collection of items to add.
' The collection can be a <b>VBA.Collection</b>, <b>ICollection</b> object,
' or an <b>Array</b>.
' @remarks The elements in the collection added by using For..Each to
' iterate each element. The elements are added to the end of the list in the
' same order as the For..Each loop produces them.
' <p>If the number of elements to be added is more than the available
' capacity in the <b>ArrayList</b>, then the capacity is increased to accommodate
' the collection of elements.</p>
' @see ICollection
' @include "..\..\Includes\ArrayListAddRange.txt"
Public Sub AddRange(ByRef Source As Variant)
    InsertRange mCount, Source
End Sub

''
' Performs a binary search for the value in the internal list.
'
' @param Value The value to search for.
' @param Comparer A comparer to perform user-defined comparison logic.
' @return The index at which the value was found, or a negative value indicating
' the value was not found.
' @remarks The search assumes the <b>ArrayList</b> is sorted.
' <p>The default comparison method requires the values to be of the same
' intrinsic Visual Basic datatype. A vbLong will fail to compare against a vbInteger,
' for instance. Objects must implement the <b>IComparable</b> interface or
' an exception will be thrown.</p>
' <p>The comparison behaviour can be overridden by supplying a custom <b>IComparer</b>
' compatible object.</P>
' <p>The search can be limited to a specific range within the list instead
' of over the entire list.</p>
' <p>If the return value is negative, then the value was not found in the list. To
' determine where the value should have been found, negate (Not) the return value.</p>
' @see IComparer
' @see IComparable
' @include "..\..\Includes\ArrayListBinarySearch.txt"
Public Function BinarySearch(ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    BinarySearch = BinarySearchEx(0, mCount, Value, Comparer)
End Function

''
' Performs a binary search for the value in the internal list.
'
' @param Index  The starting index to begin the search.
' @param Count The number of elements to search in from the startindex.
' @param Value The value to search for.
' @param Comparer A comparer to perform user-defined comparison logic.
' @return The index at which the value was found, or a negative value indicating
' the value was not found.
' @remarks The search assumes the <b>ArrayList</b> is sorted.
' <p>The default comparison method requires the values to be of the same
' intrinsic Visual Basic datatype. A vbLong will fail to compare against a vbInteger,
' for instance. Objects must implement the <b>IComparable</b> interface or
' an exception will be thrown.</p>
' <p>The comparison behaviour can be overridden by supplying a custom <b>IComparer</b>
' compatible object.</P>
' <p>The search can be limited to a specific range within the list instead
' of over the entire list.</p>
' <p>If the return value is negative, then the value was not found in the list. To
' determine where the value should have been found, negate (Not) the return value.</p>
' @see IComparer
' @see IComparable
' @include "..\..\Includes\ArrayListBinarySearch.txt"
Public Function BinarySearchEx(ByVal Index As Long, ByVal Count As Long, ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    If Index < 0 Then _
        Error.NegativeNumber "Index"
    If Count < 0 Then _
        Error.NegativeNumber "Count"
    If Index + Count > mCount Then _
        Error.InvalidOffsetLength
    
    If mCount = 0 Then
        BinarySearchEx = -1
        Exit Function
    End If
    
    If Comparer Is Nothing Then
        Set Comparer = mComparer
    End If
    
    BinarySearchEx = CorArray.BinarySearchEx(mItems, Index, Count, Value, Comparer)
End Function

''
' Returns the total number of elements in the the internal array.
'
' @return The number of elements allocated internally.
' @remarks This is not the same as <b>Count</b>. <b>Count</b> returns
' the number of items that have been added to the list.
'
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the number of allocated elements in the array.
'
' @param Value The number of elements to set the array to.
' @remarks The capacity cannot be set less than the number of items in the list.
' If the capacity is set to zero, then it is set to the default capacity (16).
' @include "..\..\Includes\ArrayListCapacity.txt"
Public Property Let Capacity(ByVal Value As Long)
    If Value < mCount Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_SmallCapacity
    
    If Value = 0 Then
        Value = DEF_CAPACITY
    End If
    
    If Value <> mCapacity Then
        mCapacity = Value
        ReDim Preserve mItems(0 To mCapacity - 1)
    End If
End Property

''
' Clears the list of all values.
'
' @remarks The internal array is cleared of values. The count is reset
' to zero, but the capacity remains the same. To lower the number of
' elements in the array, use TrimToSize or set Capacity directly.
'
Public Sub Clear()
    ReDim mItems(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Returns an ArrayList object containing a copy of the array in
' the original ArrayList.
'
' @return A copy of the original object.
' @remarks When the list of elements is cloned, elements that are
' reference types are not cloned. The cloned list will contain reference
' pointers to the same objects as the original list has.
'
Public Function Clone() As ArrayList
    Set Clone = New ArrayList
    Clone.InitClone mItems, mCount, mComparer
End Function

''
' Returns if the list contains a specific value.
'
' @param value The value to search for.
' @return A boolean value indicating the value was found in the list.
' @remarks This method performs a linear search of all elements in the list
' comparing each element with the value being searched for. If the value is
' an object, then if it implements the <b>cObject</b> interface, then the
' <b>Equals</b> method is used to test equality, otherwise the <b>Is</b> operator
' is use to compare two objects.
' <p>If there is special compare logic necessary, then a comparer
' object can be supplied. The compare object must implement the IComparer interface.</p>
'
Public Function Contains(ByRef Value As Variant) As Boolean
    Contains = IndexOf(Value) >= 0
End Function

''
' Copies all of the items to an array.
'
' @param Arr The array to copy the items to.
' @param ArrayIndex  The index to start copying items to in the destination array.
' @remarks The destination array must be large enough to hold all of the items. If an
' arrayindex is specified, then the destination array must be large enough to hold all of
' the items from that index on.
' @include "..\..\Includes\ArrayListCopyTo.txt"
Public Sub CopyTo(ByRef Arr As Variant, Optional ByRef ArrayIndex As Variant)
    If IsNullArray(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    
    CorArray.CopyEx mItems, 0, Arr, IfLong(ArrayIndex, LBound(Arr)), mCount
End Sub

''
' Copies a specified number of the items to an Array.
'
' @param Index The starting index in the list to start copying from.
' @param Arr The array to copy the items to.
' @param ArrayIndex The index in the destination array to start copying items to.
' @param Count The number of items to copy.
' @remarks The destination array must be large enough to hold the number of elements being copied.
'
Public Sub CopyToEx(ByVal Index As Long, ByRef Arr As Variant, ByVal ArrayIndex As Long, ByVal Count As Long)
    If mCount - Index < Count Then _
        Error.InvalidOffsetLength
    If IsNullArray(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    If IsMultiDimArray(Arr) Then _
        Error.Rank
        
    CorArray.CopyEx mItems, Index, Arr, ArrayIndex, Count
End Sub

''
' Returns the number of items in the list.
'
' @return The number of items in the list.
'
Public Property Get Count() As Long
    Count = mCount
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns an enumerator for an ArrayList.
'
' @param Index  The index in the list to being enumeration.
' @param Count  The number of items in the list to enumerate over.
' @return An ArrayList enumerator
' @remarks This enumerator can be used in For..Each loops. To access the optional
' parameters, the GetEnumerator must be called instead of simply passing the object
' to the For..Each. This allows for a portion of the list to be enumerated.
' @include "..\..\Includes\ArrayListGetEnumerator.txt"
Public Function GetEnumerator(Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Object
    Dim Range As ListRange
    Range = GetOptionalRange(Index, Count, 0, mCount)
    If Range.Index < 0 Then _
        Error.NegativeNumber "Index"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count > mCount Then _
        Error.InvalidOffsetLength

    Dim Ret As New ArrayListEnumerator
    Ret.Init Me, Range.Index, Range.Count
    Set GetEnumerator = Ret
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns an ArrayList object that is a windowed view into the original ArrayList.
'
' @param Index The index in the original list to start the new view.
' @param Count The number of elements in the original list to show in the view.
' @return An ArrayList that represents a subset view of the original list.
' @remarks  This windowed view is the only way to then modify the original
' ArrayList. If the original list is modified some other way, then an exception
' is thrown the next time the ranged view is accessed.
'
Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As ArrayList
    If Index < 0 Then _
        Error.NegativeNumber "Index"
    If Count < 0 Then _
        Error.NegativeNumber "Count"
    If Index + Count > mCount Then _
        Error.InvalidOffsetLength
        
    Dim Ranged As New RangedArrayList
    Ranged.Init Me, Index, Count
    Set GetRange = Ranged
End Function

''
' Returns the index of the first occurrence of the value in the list.
'
' @param value The value to search the list for.
' @param startindex The index to begin the search in the list.
' @param count The number of items in the list to search.
' @return The index of the first occurrence of value, or -1 if the value was not found.
' @remarks If there is special compare logic necessary, then a comparer
' object can be supplied. The compare object must implement the IComparer interface.
'
Public Function IndexOf(ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    Range = GetOptionalStepRange(StartIndex, Count, 0, mCount, Parameter_StartIndex)
    
    If Range.Index > mCount Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    If Range.Count < 0 Or Range.Index + Range.Count > mCount Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_Count
    
    If mCount = 0 Then
        IndexOf = -1
        Exit Function
    End If
    
    IndexOf = CorArray.IndexOf(mItems, Value, StartIndex, Count)
End Function

''
' Inserts a value into the list at the specified index.
'
' @param index The index in which the value is to be inserted.
' @param value The value to be inserted into the list.
' @remarks When a value is inserted, all items starting at the insertion point
' are moved out by one space. If capacity is reached, the list will be reallocated.
'
Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
    If Index < 0 Or Index > mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayListInsert
    
    EnsureCapacity mCount + 1
    InsertSpace Index, 1
    VariantCopyInd mItems(Index), Value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

''
' Inserts a collection of items into the list.
'
' @param index The index at which the items will be inserted.
' @param c The collection of items to be inserted into the list.
' @remarks When the items are inserted, the items starting at the
' index will be moved out a number of spaced equal to the number of
' items to be inserted. If the capacity is reached, the list will
' be reallocated.<br><br>
'
' The collection can be a <b>VBA.Collection</b>,<b>ICollection</b> object,
' or an <b>Array</b>.
'
Public Sub InsertRange(ByVal Index As Long, ByRef Source As Variant)
    If Index < 0 Or Index > mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_ArrayListInsert
    
    WriteRange Index, Source, True
    mVersion = mVersion + 1
End Sub

''
' Returns if the list is fixed-size.
'
' @return Value indicating if the list is fixed-size.
' @remarks All operations on the list can be executed as long as the number
' of items don't change, this includes the capacity of the underlying array.
'
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Returns if the list is read-only.
'
' @return Value indicating if the list is read-only.
' @remarks All operations on the list can be executed as long as the number
' of items don't change, and the items aren't changed in the list.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns an item from the list.
'
' @param index The index in the list from which to return the item.
' @return The item as the specified index.
' @remarks This property is set as the <i>Default</i> property for
' this class. Items can be accessed parentheses:<br><br>
' var = list(index)<br><br>
' The index is zero-based.
'
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    VariantCopy Item, mItems(Index)
End Property

''
' Sets an item in the list to a value.
'
' @param index The index in the list to set the value.
' @param RHS The value to set in the list.
' @remarks This property is set as the <i>Default</i> property for
' this class. Items can be accessed using parentheses:<br><br>
' list(index) = value<br><br>
' The index is zero-based.
'
Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    mItems(Index) = Value
    mVersion = mVersion + 1
End Property

''
' Sets an item in the list to a value.
'
' @param index The index in the list to set the value.
' @param RHS The value to set in the list.
' @remarks This property is set as the <i>Default</i> property for
' this class. Items can be accessed parentheses:<br><br>
' Set list(index) = value<br><br>
' The index is zero-based.
'
Public Property Set Item(ByVal Index As Long, ByRef Value As Variant)
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    Set mItems(Index) = Value
    mVersion = mVersion + 1
End Property

''
' Returns the last occurrence of a value in the list.
'
' @param Value The value to search for in the list.
' @param StartIndex The index to start searching from.
' @param Count The number of items to be searched.
' @return The index of the last occurrence of the value, or -1 if not found.
' @remarks The search is performed from the end of the list towards the
' beginning. If a startindex
'
Public Function LastIndexOf(ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As Long
    Dim Range As ListRange
    Range.Index = IfLong(StartIndex, mCount - 1)
    Range.Count = IfLong(Count, Range.Index + 1)
    
    If Range.Index < 0 Then _
        Error.NegativeNumber "StartIndex"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index >= mCount Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_LargerThanCollection
    If Range.Index - Range.Count + 1 < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_LargerThanCollection
    
    If mCount = 0 Then
        LastIndexOf = -1
        Exit Function
    End If
    
    LastIndexOf = CorArray.LastIndexOf(mItems, Value, StartIndex, Count)
End Function

''
' Returns an enumerator
'
' @return The enumerator
' @see IEnumerator
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Removes a value from the list.
'
' @param value The value to be removed.
' @remarks If the value is an object, then it must be the same instance of
' the object to be removed. This behaviour can be overriden with a custom comparer.
' <p>When removing an item, the remaining items maintain there current order
' relative to the other items.</p>
'
Public Sub Remove(ByRef Value As Variant)
    Dim Index As Long
    Index = IndexOf(Value)
    
    If Index >= 0 Then
        RemoveAt Index
    End If
End Sub

''
' Removes a value from the list at a specific index.
'
' @param index The index in the list to remove the item.<br>
' The index is zero-based.
' @remarks When removing an item, the remaining items maintain there current order
' relative to the other items.
'
Public Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    mItems(Index) = Empty
    RemoveSpace Index, 1
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

''
' Remove a set of items from the list.
'
' @param index The index to start removing items.
' @param count The number of items to remove.
' @remarks The index is zero-based.
'
Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    If Index < 0 Then _
        Error.NegativeNumber "Index"
    If Count < 0 Then _
        Error.NegativeNumber "Count"
    If Index + Count > mCount Then _
        Error.InvalidOffsetLength

    Dim i As Long
    For i = Index To Index + Count - 1
        mItems(i) = Empty
    Next
    
    RemoveSpace Index, Count
    mCount = mCount - Count
    mVersion = mVersion + 1
End Sub

''
' Reverses the list of items in the list.
'
' @param index The index to start reversing items.
' @param count The number of items to include in the reverse.
' @remarks If no parameters are supplied, then then
' entire list will be reversed.
'
Public Sub Reverse(Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim Range As ListRange
    Range = GetOptionalRange(Index, Count, 0, mCount)

    If Range.Index < 0 Then _
        Error.NegativeNumber "Index"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count > mCount Then _
        Error.InvalidOffsetLength


    CorArray.Reverse mItems, Range.Index, Range.Count
    mVersion = mVersion + 1
End Sub

''
' Sets the items in the list to a collection of items.
'
' @param Index The index at which to start setting items.
' @param Source A collection of items to set in the list.
' @remarks The number of items in the list is not changed, nor is the size
' of the list increased. Only existing items in the list can be set to
' the items in the collection. This basically overlays new items onto
' existing items.
'
Public Sub SetRange(ByVal Index As Long, ByRef Source As Variant)
    WriteRange Index, Source, False
    mVersion = mVersion + 1
End Sub

''
' Sorts the items in the list.
'
' @param Index The index to start sorting the items.
' @param Count The number of items to include in the sort.
' @param Comparer A custom comparer used to compare items in the list.
' @remarks The list is sorted in ascending order by default. By providing
' a custom comaprer, the list can be sorted in special ways, such as descending order.
'
Public Sub Sort(Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Comparer As IComparer)
    Dim Range As ListRange
    Range = GetOptionalRange(Index, Count, 0, mCount)

    If Range.Index < 0 Then _
        Error.NegativeNumber "Index"
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count > mCount Then _
        Error.InvalidOffsetLength


    If Comparer Is Nothing Then
        Set Comparer = mComparer
    End If
    
    CorArray.SortEx mItems, Range.Index, Range.Count, Comparer
    mVersion = mVersion + 1
End Sub

''
' Returns an array of the items in the list.
'
' @param arraytype The specific array type to return.
' @return An array of items in the list.
' @remarks <p>If an array type other than variant is specified, then the
' list attempts to convert each item in the list to the specific array type
' requested. If an item cannot be converted, then an exception is thrown.</p>
' <p>No items will be narrowed in the conversion. For example, a Double will
' not be converted to a Long. An exception will be thrown.</p>
'
Public Function ToArray(Optional ByVal ArrayType As VbVarType = vbVariant) As Variant
    Dim Ret As Variant
    Ret = CorArray.CreateInstance(ArrayType, mCount)
    
    If mCount > 0 Then
        CorArray.CopyEx mItems, 0, Ret, 0, mCount
    End If
    
    ToArray = Ret
End Function

''
' Sets the capacity to the number of items in the list.
'
' @remarks This will decrease the size of the internal array. This can save
' memory if you know that no more items will be added to the list.
'
Public Sub TrimToSize()
    Capacity = mCount
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

'Public Property Get Version() As Long
'    Version = mVersion
'End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub InitClone(ByRef Arr() As Variant, ByVal Count As Long, ByVal Comparer As IComparer)
    mItems = Arr
    mCount = Count
    mCapacity = UBound(Arr) + 1
    Set mComparer = Comparer
End Sub

Friend Sub Init(ByVal Comparer As IComparer, ByRef Source As Variant)
    Set mComparer = Comparer
    If Not IsMissing(Source) Then
        AddRange Source
    End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Private Sub InsertSpace(ByVal Index As Long, ByVal Size As Long)
    If Index < mCount Then
        CopyMemory mItems(Index + Size), mItems(Index), (mCount - Index) * SizeOfVariant
        ZeroMemory mItems(Index), Size * SizeOfVariant
    End If
End Sub

Private Sub RemoveSpace(ByVal Index As Long, ByVal Size As Long)
    If Index < mCount - 1 Then
        CopyMemory mItems(Index), mItems(Index + Size), (mCount - Index - Size) * SizeOfVariant
        ZeroMemory mItems(mCount - Size), Size * SizeOfVariant
    End If
End Sub

Private Sub WriteRange(ByVal Index As Long, ByRef Source As Variant, ByVal Insert As Boolean)
    If IsArray(Source) Then
        If IsNullArray(Source) Then _
            Error.ArgumentNull "Source", ArgumentNull_Array
        If IsMultiDimArray(Source) Then _
            Error.Rank
            
        WriteArray Index, Source, Insert
    ElseIf IsObject(Source) Then
        If Source Is Nothing Then _
            Error.ArgumentNull "Source", ArgumentNull_Collection
            
        If TypeOf Source Is Collection Then
            WriteVBCollection Index, Source, Insert
        ElseIf TypeOf Source Is ICollection Then
            WriteICollection Index, Source, Insert
        Else
            Throw Cor.NewInvalidCastException("An ICollection or VBA.Collection object is required.")
        End If
    Else
        Throw Cor.NewInvalidCastException("An Array, ICollection, or VBA.Collection object is required.")
    End If
End Sub

Private Sub WriteArray(ByVal Index As Long, ByRef Source As Variant, ByVal Insert As Boolean)
    Dim SizeOfArray As Long
    SizeOfArray = UBound(Source) - LBound(Source) + 1
    
    WriteCollection Index, SizeOfArray, Source, Insert
End Sub

Private Sub WriteVBCollection(ByVal Index As Long, ByVal Source As Collection, ByVal Insert As Boolean)
    WriteCollection Index, Source.Count, Source, Insert
End Sub

Private Sub WriteICollection(ByVal Index As Long, ByVal Source As ICollection, ByVal Insert As Boolean)
    WriteCollection Index, Source.Count, Source, Insert
End Sub

Private Sub WriteCollection(ByVal Index As Long, ByVal SizeOfCollection As Long, ByRef Source As Variant, ByVal Insert As Boolean)
    If Insert Then
        EnsureCapacity mCount + SizeOfCollection
        InsertSpace Index, SizeOfCollection
        mCount = mCount + SizeOfCollection
    End If
        
    Dim Value As Variant
    For Each Value In Source
        Helper.MoveVariant mItems(Index), Value
        Index = Index + 1
    Next Value
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Capacity = DEF_CAPACITY
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    On Error GoTo errTrap
    With PropBag
        mCount = .ReadProperty(PROP_COUNT, 0)
        EnsureCapacity .ReadProperty(PROP_CAPACITY, DEF_CAPACITY)
        
        Dim i As Long
        For i = 0 To mCount - 1
            ' When persisting data within a Class_WriteProperties event, it seems that
            ' object values of Nothing and variant values of Empty do not get persisted
            ' correctly, or at all, except the key name. Even when reading with a default
            ' value of Empty, it returns an empty string "". So now we maintain a flag
            ' with each value to determine the correct value.
            Select Case .ReadProperty(PROP_SUBTYPE & i, SUBTYPE_NORMAL)
                Case SUBTYPE_NORMAL
                    Helper.MoveVariant mItems(i), .ReadProperty(PROP_ITEMPREFIX & i)
                
                Case SUBTYPE_EMPTY
                    mItems(i) = Empty
                
                Case SUBTYPE_NOTHING
                    Set mItems(i) = Nothing
            End Select
        Next i
        
        If .ReadProperty(PROP_DEFAULTCOMPARER, False) Then
            Set mComparer = Comparer.Default
        Else
            Set mComparer = .ReadProperty(PROP_COMPARER, Nothing)
        End If
    End With
    Exit Sub
    
errTrap:
    Throw Cor.NewSerializationException(Err.Description)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    On Error GoTo errTrap
    With PropBag
        Dim i As Long
        For i = 0 To mCount - 1
            ' When persisting data within a Class_WriteProperties event, it seems that
            ' object values of Nothing and variant values of Empty do not get persisted
            ' correctly, or at all, except the key name. Even when reading with a default
            ' value of Empty, it returns an empty string "". So now we maintain a flag
            ' with each value to determine the correct value.
            Select Case VarType(mItems(i))
                Case vbEmpty
                    PropBag.WriteProperty PROP_SUBTYPE & i, SUBTYPE_EMPTY
                
                Case vbObject
                    If mItems(i) Is Nothing Then
                        PropBag.WriteProperty PROP_SUBTYPE & i, SUBTYPE_NOTHING
                    Else
                        PropBag.WriteProperty PROP_ITEMPREFIX & i, mItems(i)
                    End If
                
                Case Else
                    PropBag.WriteProperty PROP_ITEMPREFIX & i, mItems(i)
            
            End Select
        Next i
        
        .WriteProperty PROP_COUNT, mCount
        .WriteProperty PROP_CAPACITY, mCapacity
        If mComparer Is Comparer.Default Then
            .WriteProperty PROP_DEFAULTCOMPARER, True
        Else
            .WriteProperty PROP_COMPARER, mComparer, Nothing
        End If
    End With
    Exit Sub
    
errTrap:
    Throw Cor.NewSerializationException(Err.Description)
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IList Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IList_Add(Value As Variant) As Long
    IList_Add = Add(Value)
End Function

Private Sub IList_Clear()
    Clear
End Sub

Private Function IList_Contains(Value As Variant) As Boolean
    IList_Contains = Contains(Value)
End Function

Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub

Private Property Get IList_Count() As Long
    IList_Count = Count
End Property

Private Function IList_GetEnumerator() As IEnumerator
    Set IList_GetEnumerator = GetEnumerator
End Function

Private Function IList_IndexOf(Value As Variant) As Long
    IList_IndexOf = IndexOf(Value)
End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant)
    Insert Index, Value
End Sub

Private Property Get IList_IsFixedSize() As Boolean
    IList_IsFixedSize = IsFixedSize
End Property

Private Property Get IList_IsReadOnly() As Boolean
    IList_IsReadOnly = IsReadOnly
End Property

Private Property Set IList_Item(ByVal Index As Long, Value As Variant)
    Set Item(Index) = Value
End Property

Private Property Let IList_Item(ByVal Index As Long, Value As Variant)
    Item(Index) = Value
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
    Helper.MoveVariant IList_Item, Item(Index)
End Property

Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = NewEnum
End Function

Private Sub IList_Remove(Value As Variant)
    Remove Value
End Sub

Private Sub IList_RemoveAt(ByVal Index As Long)
    RemoveAt Index
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IVersionableCollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IVersionableCollection_CurrentVersion() As Long
    IVersionableCollection_CurrentVersion = mVersion
End Property
