VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedList"
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
' Module: SortedList
'

''
' Provides a means to maintain a list of sorted items based on an associated key.
'
' @remarks The <b>SortedList</b> maintains sort order for every key/value pair
' added to the list. The list is expanded as needed with the addition of new items.
' If a key already exists in the list, an exception is thrown.
'
' @see Constructors
' @see IDictionary
' @see ICollection
' @see IEnumerable
'
'@Folder("CorLib.System.Collections")
Option Explicit
Implements IObject
Implements IEnumerable
Implements ICollection
Implements IDictionary
Implements IVersionable

Private Const DefaultCapacity   As Long = 16
Private Const PropCount         As String = "Count"
Private Const PropCapacity      As String = "Capacity"
Private Const PropComparer      As String = "Comparer"


Private mKeys()     As Variant
Private mValues()   As Variant
Private mComparer   As IComparer
Private mCapacity   As Long
Private mCount      As Long
Private mVersion    As Long


''
' Adds an element with the specified key and value to a <a href="SortedList.html">SortedList</a> object..
'
' @param key The value to used to maintain a sorted order.
' @param value The value associated with the key.
' @remark The insertion point is determined based on the comparer selected, either explicitly or by default, when the SortedList object was created.
' <p>You can also use the <a href="SortedList.Let.Item.html">Item</a> property to add new elements by setting the value of a key that does not exist in the <a href="SortedList.html">SortedList</a> object
' (for example, myCollection("myNonexistentKey") = myValue). However, if the specified key already exists in the <a href="SortedList.html">SortedList</a>, setting the <a href="SortedList.Let.Item.html">Item<a/> property overwrites the old value.
' In contrast, the <b>Add</b> method does not modify existing elements.</p>
' <p>The elements of a <a href="SortedList.html">SortedList</a> object are sorted by the keys either according to a specific <a href="IComparer.html">IComparer</a> implementation
' specified when the <a href="SortedList.html">SortedList</a> is created.</p>
' @see #Item
' @see IComparer
' @include "..\DocIncludes\SortedList.Add.html"
Public Sub Add(ByRef Key As Variant, ByRef Value As Variant)
    Dim Index As Long
    
    If mCount > 0 Then
        Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
        
        If Index >= 0 Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_AddingDuplicate_Key, GetKey(Index), Key))
            
        Index = Not Index
    End If
    
    Insert Index, Key, Value
End Sub


''
' Gets the capacity of the <a href="SortedList.html">SortedList</a> object.
'
' @return The number of elements that the <a href="SortedList.html">SortedList</a> object can contain.
' @remarks <b>Capacity</b> is the number of elements that the <a href="SortedList.html">SortedList</a> object can store. <b>Count</b> is the number of elements that are actually in the <a href="SortedList.html">SortedList</a>.
' <p><b>Capacity</b> is always greater than or equal to <a href="SortedList.Get.Count.html">Count</a>. If <a href="SortedList.Get.Count.html">Count</a> exceeds <b>Capacity</b> while adding elements,
' the capacity is automatically increased by reallocating the internal array before copying the old elements and adding the new elements.</p>
' <p>The capacity can be decreased by calling <a href="SortedList.TrimToSize.html">TrimToSize</a> or by setting the <b>Capacity</b> property explicitly.
' When the value of <b>Capacity</b> is set explicitly, the internal array is also reallocated to accommodate the specified capacity.</p>
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Gets the capacity of the <a href="SortedList.html">SortedList</a> object.
'
' @param Value The number of elements the <a href="SortedList.html">SortedList</a> object can contain.
' @remarks <b>Capacity</b> is the number of elements that the <a href="SortedList.html">SortedList</a> object can store. <b>Count</b> is the number of elements that are actually in the <a href="SortedList.html">SortedList</a>.
' <p><b>Capacity</b> is always greater than or equal to <a href="SortedList.Get.Count.html">Count</a>. If <a href="SortedList.Get.Count.html">Count</a> exceeds <b>Capacity</b> while adding elements,
' the capacity is automatically increased by reallocating the internal array before copying the old elements and adding the new elements.</p>
' <p>The capacity can be decreased by calling <a href="SortedList.TrimToSize.html">TrimToSize</a> or by setting the <b>Capacity</b> property explicitly.
' When the value of <b>Capacity</b> is set explicitly, the internal array is also reallocated to accommodate the specified capacity.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>The value assigned is less than the current number of elements in the <a href="SortedList.html">SortedList</a> object.</td>
'   </tr>
' </table>
Public Property Let Capacity(ByVal Value As Long)
    If Value < mCount Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_SmallCapacity
    
    If Value <> mCapacity Then
        If Value > 0 Then
            ReDim Preserve mKeys(0 To Value - 1)
            ReDim Preserve mValues(0 To Value - 1)
        Else
            mKeys = Array()
            mValues = Array()
        End If
    End If
    
    mCapacity = Value
End Property

''
' Removes all of the items from the <a href="SortedList.html">SortedList</a> object.
'
' @remarks <a href="SortedList.Get.Count.html">Count</a> is set to zero and elements are cleared from the collection.
' <p><a href="SortedList.Get.Capacity.html">Capacity</a> remains unchanged. To reset the capacity of the <a href="SortedList.html">SortedList</a> object,
' call <a href="SortedList.TrimToSize.html">TrimToSize</a> or set the <a href="SortedList.Let.Capacity.html">Capacity</a> property directly.</p>
Public Sub Clear()
    ReDim mKeys(0 To mCapacity - 1)
    ReDim mValues(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Returns a cloned version of this instance.
'
' @remarks Reference types are not cloned. If there are object types
' in the list, then a new reference to the object is created, but the
' object itself is not cloned.
Public Function Clone() As SortedList
    Set Clone = New SortedList
    Clone.InitClone mKeys, mValues, mCount, mComparer
End Function

''
' Determines whether a <a href="SortedList.html">SortedList</a> object contains a specific key.
'
' @param Key The key to locate in the <a href="SortedList.html">SortedList</a> object.
' @return <b>True</b> if the <a href="SortedList.html">SortedList</a> object contains an element with specified <i>Key</i>; otherwise <b>False</b>.
' @remarks The elements of a <a href="SortedList.html">SortedList</a> object are sorted by the keys either according to a specific <a href="IComparer.html">IComparer</a>
' implementation specified when the <a href="SortedList.html">SortedList</a> is created.
' <p>This method behaves exactly as the <a href="SortedList.ContainsKey.html">ContainsKey</a> method.</p>
' <p>This method uses a binary search algorithm; therefore, this method is an O(log <i>n</i>) operation, where <i>n</i> is <a href="SortedList.Get.Count.html">Count</a>.</p>
' @see #ContainsKey
' @see #ContainsValue
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The comparer throws an exception.</td>
'   </tr>
' </table>
' @include "..\DocIncludes\SortedList.Contains.html"
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = ContainsKey(Key)
End Function

''
' Determines whether a <a href="SortedList.html">SortedList</a> object contains a specific key.
'
' @param Key The key to locate in the <a href="SortedList.html">SortedList</a> object.
' @return Indication of the key existing in the list.
' @remarks The elements of a <a href="SortedList.html">SortedList</a> object are sorted by the keys either according to a specific <a href="IComparer.html">IComparer</a>
' implementation specified when the <a href="SortedList.html">SortedList</a> is created.
' <p>This method behaves exactly as the <a href="SortedList.Contains.html">Contains</a>.</p>
' <p>This method uses a binary search algorithm; therefore, this method is an O(log <i>n</i>) operation, where <i>n</i> is <a href="SortedList.Get.Count.html">Count</a>.</p>
' @see #Contains
' @see #ContainsValue
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The comparer throws an exception.</td>
'   </tr>
' </table>
' @include "..\DocIncludes\SortedList.Contains.html"
Public Function ContainsKey(ByRef Key As Variant) As Boolean
    ContainsKey = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer) >= 0
End Function

''
' Determines whether a <a href="SortedList.html">SortedList</a> object contains a specific value.
'
' @param value The value to locate for in the <a href="SortedList.html">SortedList</a> object.
' @return <b>True</b> if the <a href="SortedList.html">SortedList</a> object contains an element with the specified <i>Value</i>; otherwise, <b>False</b>.
' @remarks This method performs a linear search; therefore, the average execution time is proportional to <a href="SortedList.Get.Count.html">Count</a>.
' That is, this method is an O(<i>n</i>) operation, where <i>n</i> is <a href="SortedList.Get.Count.html">Count</a>.
' @see Contains
' @see ContainsKey
' @include "..\DocIncludes\SortedList.Contains.html"
Public Function ContainsValue(ByRef Value As Variant) As Boolean
    ContainsValue = CorArray.IndexOf(mValues, Value, 0, mCount) >= 0
End Function

''
' Copies <a href="SortedList.html">SortedList</a> elements to a one-dimensional array, starting at the specified index in the array.
'
' @param Arr The one-dimensional array that is the destination of the <a href="DictionaryEntry.html">DictionaryEntry</a> objects copied from <a href="SortedList.html">SortedList</a>.
' @param ArrayIndex The starting index in <i>Arr</i> to begin copying to.
' @remarks The key/value pairs are copied to the array in the same order in which the enumerator iterates through the <a href="SortedList.html">SortedList</a> object.
' <p>To copy only the keys in the <a href="SortedList.html">SortedList</a>, use <code>SortedList.Keys.CopyTo</code>.</p>
' <p>To copy only the values in the <a href="SortedList.html">SortedList</a>, use <code>SortedList.Values.CopyTo</code>.</p>
' @see DictionaryEntry
' @include "..\DocIncludes\SortedList.CopyTo.html"
Public Sub CopyTo(ByRef Arr As Variant, ByVal ArrayIndex As Long)
    If CorArray.IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    If CorArray.Rank(Arr) <> 1 Then _
        Error.Argument Arg_RankMultiDimNotSupported, "Arr"
    If ArrayIndex < LBound(Arr) Then _
        Error.ArgumentOutOfRange "ArrayIndex", ArgumentOutOfRange_LBound
    If ArrayIndex + mCount - 1 > UBound(Arr) Then _
        Error.Argument Arg_ArrayPlusOffTooSmall
    
    On Error GoTo Catch
    
    Dim i As Long
    For i = 0 To mCount - 1
        Set Arr(ArrayIndex + i) = Cor.NewDictionaryEntry(mKeys(i), mValues(i))
    Next i
    
    Exit Sub
    
Catch:
    If Err.Number = vbTypeMismatch Then
        Throw Cor.NewInvalidCastException(Environment.GetResourceString(InvalidCast_StoreArrayElement))
    Else
        Throw Err
    End If
End Sub

''
' Gets the number of elements contained in a <a href="SortedList.html">SortedList</a> object.
'
' @return The number of elements in the <a href="SortedList.html">SortedList</a> object.
' @remarks <b>Capacity</b> is the number of elements that the <a href="SortedList.html">SortedList</a> object can store. <b>Count</b> is the number of elements that are actually in the <a href="SortedList.html">SortedList</a>.
' <p><b>Capacity</b> is always greater than or equal to <a href="SortedList.Get.Count.html">Count</a>. If <a href="SortedList.Get.Count.html">Count</a> exceeds <b>Capacity</b> while adding elements,
' the capacity is automatically increased by reallocating the internal array before copying the old elements and adding the new elements.</p>
' @see Capacity
Public Property Get Count() As Long
    Count = mCount
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Gets the value at the specified index of a <a href="SortedList.html">SortedList</a> object.
'
' @param Index The zero-based index of the value to get.
' @return The value at the specified index in the <a href="SortedList.html">SortedList</a> object.
' @remarks The index sequence is based on the sort sequence. When an element is added, it is inserted into
' <a href="SortedList.html">SortedList</a> in the correct sort order, and the indexing adjusts accordingly.
' When an element is removed, the indexing also adjusts accordingly. Therefore, the index of a specific key/value pair
' might change as elements are added or removed from the <a href="SortedList.html">SortedList</a> object.
' @see #IndexOfKey
' @see #IndexOfValue
' @include "..\DocIncludes\SortedList.GetByIndex.html"
Public Function GetByIndex(ByVal Index As Long) As Variant
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    VariantCopy GetByIndex, mValues(Index)
End Function

''
' Returns an enumerator that can be used to iterate through the key/value pairs.
'
' @return An IEnumerator object to be use for iterating the key/value pairs.
'
Public Function GetEnumerator() As IEnumerator
    Dim Ret As New SortedListEnumerator
    Ret.Init Me, mKeys, mValues, slEntries
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
' Returns the key at the specified index in the list.
'
' @param Index The index into the list from which to retrieve the key.
' @return The key at the specified index.
' @remarks Since the list maintains sort order, when an item is
' added or removed, there is no garauntee that the same key will be
' found at the same index in future calls.
' @include "..\DocIncludes\SortedList.GetKey.html"
Public Function GetKey(ByVal Index As Long) As Variant
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
        
    VariantCopy GetKey, mKeys(Index)
End Function

''
' Gets the keys in a <a href="SortedList.html">SortedList</a> object.
'
' @return An <a href="IList.html">IList</a> object containing the keys in the <a href="SortedList.html">SortedList</a> object.
' @remarks The returned <a href="IList.html">IList</a> object is a read-only view of the keys of the <a href="SortedList.html">SortedList</a> object.
' Modifications made to the underlying <a href="SortedList.html">SortedList</a> are immediately reflected in the <a href="IList.html">IList</a>.
' <p>The elements of the returned <a href="IList.html">IList</a> are sorted in the same order as the keys of the <a href="SortedList.html">SortedList</a>.</p>
' <p>This method is similar to the <a href="SortedList.Get.Keys.html">Keys</a> property, but returns an <a href="IList.html">IList</a> object instead of an <a href="ICollection.html">ICollection</a> object.</p>
' @see IList
' @see #GetValueList
' @see #Keys
' @include "..\DocIncludes\SortedList.GetKeyList.html"
Public Function GetKeyList() As IList
    Set GetKeyList = NewSortedKeyList
End Function

''
' Gets the values in a <a href="SortedList.html">SortedList</a> object.
'
' @return An <a href="IList.html">IList</a> object containing the values in the <a href="SortedList.html">SortedList</a> object.
' @remarks The returned <a href="IList.html">IList</a> object is a read-only view of the values of the <a href="SortedList.html">SortedList</a> object.
' Modifications made to the underlying <a href="SortedList.html">SortedList</a> are immediately reflected in the <a href="IList.html">IList</a>.
' <p>The elements of the returned <a href="IList.html">IList</a> are sorted in the same order as the values of the <a href="SortedList.html">SortedList</a>.</p>
' <p>This method is similar to the <a href="SortedList.Get.Values.html">Values</a> property, but returns an <a href="IList.html">IList</a> object instead of an <a href="ICollection.html">ICollection</a> object.</p>
' @see IList
' @see #GetKeyList
' @see #Values
' @include "..\DocIncludes\SortedList.GetValueList.html"
Public Function GetValueList() As IList
    Set GetValueList = NewSortedValueList
End Function

''
' Returns the zero-based index of the specified key in a <a href="SortedList.html">SortedList</a> object.
'
' @param Key The key to locate in the <a href="SortedList.html">SortedList</a> object.
' @return The zero-based index of the <i>Key</i> parameter, if <i>Key</i> is found in the <a href="SortedList.html">SortedList</a> object; otherwise, -1.
' @remarks The elements of a <a href="SortedList.html">SortedList</a> object are sorted by the keys either according to a specific
' <a href="IComparer.html">IComparer</a> implementation specified when the <a href="SortedList.html">SortedList</a> is created, or
' the default <a href="ComparerStatic.Get.Defaultr.html">Comparer</a>.
' <p>The index sequence is based on the sort sequence. When an element is added, it is inserted into <a href="SortedList.html">SortedList</a> in the correct sort order,
' and the indexing adjusts accordingly. When an element is removed, the indexing also adjusts accordingly. Therefore, the index of a specific key/value pair might change
' as elements are added or removed from the <a href="SortedList.html">SortedList</a>.</p>
' <p>This method uses a binary search algorithm; therefore, this method is an O(log <i>n</i>) operation, where <i>n</i> is <a href="SortedList.Get.Count.html">Count</a>.</p>
' @see #ContainsKey
' @see #IndexOfValue
' @include "..\DocIncludes\SortedList.IndexOfKey.html"
Public Function IndexOfKey(ByRef Key As Variant) As Long
    Dim Index As Long
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index < 0 Then
        Index = -1
    End If
    
    IndexOfKey = Index
End Function

''
' Returns the zero-based index of the first occurrence of the specified value in a <a href="SortedList.html">SortedList</a> object.
'
' @param Value The value to locate in the <a href="SortedList.html">SortedList</a> object. The value can be <b>Nothing</b>.
' @return The zero-based index of the first occurrence of the value parameter, if value is found in the <a href="SortedList.html">SortedList</a> object; otherwise, -1.
' @remarks The index sequence is based on the sort sequence. When an element is added, it is inserted into <a href="SortedList.html">SortedList</a> in the correct sort order,
' and the indexing adjusts accordingly. When an element is removed, the indexing also adjusts accordingly. Therefore, the index of a specific
' key/value pair might change as elements are added or removed from the <a href="SortedList.html">SortedList</a> object.
' <p>The values of the elements of the <a href="SortedList.html">SortedList</a> are compared to the specified value using the <a href="ObjectStatic.Equals.html">Equals</a> method.</p>
' <p>This method uses a linear search; therefore, this method is an O(<i>n</i>) operation, where <i>n</i> is <a href="SortedList.Get.Count.html">Count</a>.</p>
' @see #ContainsValue
' @see #IndexOfKey
' @include "..\DocIncludes\SortedList.IndexOfValue.html"
Public Function IndexOfValue(ByRef Value As Variant) As Long
    IndexOfValue = CorArray.IndexOf(mValues, Value, 0, mCount)
End Function

''
' Gets a value indicating whether a <a href="SortedList.html">SortedList</a> object has a fixed size.
'
' @return <b>True</b> if the <a href="SortedList.html">SortedList</a> object has a fixed size; otherwise, <b>False</b>. The default is <b>False</b>.
' @remarks A collection with a fixed size does not allow the addition or removal of elements after the collection is created, but does allow the modification of existing elements.
' <p>A collection with a fixed size is simply a collection with a wrapper that prevents adding and removing elements; therefore, if changes are made to the underlying collection, including the addition or removal of elements, the fixed-size collection reflects those changes.</p>
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Gets a value indicating whether a <a href="SortedList.html">SortedList</a> object is read-only.
'
' @return <b>True</b> if the <a href="SortedList.html">SortedList</a> object is read-only; otherwise, <b>False</b>. The default is <b>False</b>.
' @remarks A collection that is read-only does not allow the addition, removal, or modification of elements after the collection is created.
' <p>A collection that is read-only is simply a collection with a wrapper that prevents modifying the collection; therefore, if changes are made to the underlying collection, the read-only collection reflects those changes.</p>
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Gets the value associated with a specific key in a <a href="SortedList.html">SortedList</a> object.
'
' @param Key The key associated with the value to get.
' @return The value associated with the <i>Key</i> parameter in the <a href="SortedList.html">SortedList</a> object, if key is found; otherwise, <b>Empty</b>.
' @remarks You can use the Item property to access a specific element in a collection by specifying the following syntax: <code>myCollection(key)</code>.
' <p>You can also use this property to add new elements by setting the value of a key that does not exist in the <a href="SortedList.html">SortedList</a> object (for example, <code>myCollection("myNonexistentKey") = myValue</code>).
' However, if the specified key already exists in the <a href="SortedList.html">SortedList</a>, setting the <b>Item</b> property overwrites the old value.
' In contrast, the <a href="SortedList.Add.html">Add</a> method does not modify existing elements.</p>
' <p>The elements of a <a href="SortedList.html">SortedList</a> are sorted by the keys either according to a specific <a href="IComparer.html">IComparer</a> implementation
' specified when the <a href="SortedList.html">SortedList</a> is created, or by the default <a href="ComparerStatic.Get.Default.html">Comparer</a>.</p>
' @see #Add
' @see #Contains
' @see #ContainsKey
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The comparer throws an exception.</td>
'   </tr>
' </table>
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Index As Long
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index >= 0 Then
        VariantCopy Item, mValues(Index)
    End If
End Property

''
' Sets the value associated with a specific key in a <a href="SortedList.html">SortedList</a> object.
'
' @param Key The key associated with the value to set.
' @param Value The value associated with the <i>Key</i> parameter in the <a href="SortedList.html">SortedList</a> object.
' @remarks You can use the Item property to access a specific element in a collection by specifying the following syntax: <code>myCollection(key)</code>.
' <p>You can also use this property to add new elements by setting the value of a key that does not exist in the <a href="SortedList.html">SortedList</a> object (for example, <code>myCollection("myNonexistentKey") = myValue</code>).
' However, if the specified key already exists in the <a href="SortedList.html">SortedList</a>, setting the <b>Item</b> property overwrites the old value.
' In contrast, the <a href="SortedList.Add.html">Add</a> method does not modify existing elements.</p>
' <p>The elements of a <a href="SortedList.html">SortedList</a> are sorted by the keys either according to a specific <a href="IComparer.html">IComparer</a> implementation
' specified when the <a href="SortedList.html">SortedList</a> is created, or by the default <a href="ComparerStatic.Get.Default.html">Comparer</a>.</p>
' @see #Add
' @see #Contains
' @see #ContainsKey
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The property is set and the <a href="SortedList.html">SortedList</a> object is read-only.<br/>
'           -or-<br/>
'           The property is set, <i>Key</i> does not exist in the collection, and the <a href="SortedList.html">SortedList</a> has a fixed size.
'        </td>
'    </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The comparer throws an exception.</td>
'   </tr>
' </table>
Public Property Let Item(ByRef Key As Variant, ByRef Value As Variant)
    Dim Index As Long
    
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index >= 0 Then
        mValues(Index) = Value
        mVersion = mVersion + 1
    Else
        Insert Not Index, Key, Value
    End If
End Property

''
' Sets the value associated with a specific key in a <a href="SortedList.html">SortedList</a> object.
'
' @param Key The key associated with the value to set.
' @param Value The value associated with the <i>Key</i> parameter in the <a href="SortedList.html">SortedList</a> object.
' @remarks You can use the Item property to access a specific element in a collection by specifying the following syntax: <code>myCollection(key)</code>.
' <p>You can also use this property to add new elements by setting the value of a key that does not exist in the <a href="SortedList.html">SortedList</a> object (for example, <code>myCollection("myNonexistentKey") = myValue</code>).
' However, if the specified key already exists in the <a href="SortedList.html">SortedList</a>, setting the <b>Item</b> property overwrites the old value.
' In contrast, the <a href="SortedList.Add.html">Add</a> method does not modify existing elements.</p>
' <p>The elements of a <a href="SortedList.html">SortedList</a> are sorted by the keys either according to a specific <a href="IComparer.html">IComparer</a> implementation
' specified when the <a href="SortedList.html">SortedList</a> is created, or by the default <a href="ComparerStatic.Get.Default.html">Comparer</a>.</p>
' @see #Add
' @see #Contains
' @see #ContainsKey
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The property is set and the <a href="SortedList.html">SortedList</a> object is read-only.<br/>
'           -or-<br/>
'           The property is set, <i>Key</i> does not exist in the collection, and the <a href="SortedList.html">SortedList</a> has a fixed size.
'        </td>
'    </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The comparer throws an exception.</td>
'   </tr>
' </table>
Public Property Set Item(ByRef Key As Variant, ByRef Value As Variant)
    Dim Index As Long
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index >= 0 Then
        Set mValues(Index) = Value
        mVersion = mVersion + 1
    Else
        Insert Not Index, Key, Value
    End If
End Property

''
' Gets the keys in the <a href="SortedList.html">SortedList</a> object.
'
' @return An ICollection object containing the keys in the <a href="SortedList.html">SortedList</a> object.
' @remarks The <a href="ICollection.html">ICollection</a> object is a read-only view of the keys of the <a href="SortedList.html">SortedList</a> object.
' Modifications made to the underlying <a href="SortedList.html">SortedList</a> are immediately reflected in the <a href="ICollection.html">ICollection</a>.
' <p>The elements of the <a href="ICollection.html">ICollection</a> are sorted in the same order as the keys of the <a href="SortedList.html">SortedList</a>.</p>
' <p>This property is similar to the <a href="SortedKey.GetKeyList.html">GetKeyList</a> method,
' but returns an <a href="ICollection.html">ICollection</a> object instead of an <a href="IList.html">IList</a> object.</p>
' @see ICollection
' @see #Values
' @see #GetKeyList
Public Property Get Keys() As ICollection
    Set Keys = NewSortedKeyList
End Property

''
' Returns a For..Each compatible enumerator.
'
' @return The enumerator used to iterate the <a href="SortedList.html">SortedList</a> object.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Removes the element with the specified key from a <a href="SortedList.html">SortedList</a> object.
'
' @param Key The key of the element to remove.
' @remarks If the <a href="SortedList.html">SortedList</a> object does not contain an element with the specified key, the <a href="SortedList.html">SortedList</a> remains unchanged. No exception is thrown.
' @see #RemoveAt
' @include "..\DocIncludes\SortedList.Remove.html"
Public Sub Remove(ByRef Key As Variant)
    If mCount = 0 Then
        Exit Sub
    End If
    
    Dim Index As Long
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index >= 0 Then
        RemoveAt Index
    End If
End Sub

''
' Removes the element at the specified index of a <a href="SortedList.html">SortedList</a> object.
'
' @param Index The zero-based index of the element to remove.
' @remarks The index sequence is based on the sort sequence. When an element is added, it is inserted into <a href="SortedList.html">SortedList</a> in the correct sort order,
' and the indexing adjusts accordingly. When an element is removed, the indexing also adjusts accordingly. Therefore, the index of a specific key/value pair might change as elements are added or removed from the <a href="SortedList.html">SortedList</a> object.
' @see #Remove
' @include "..\DocIncludes\SortedList.RemoveAt.html"
Public Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    mKeys(Index) = Empty
    mValues(Index) = Empty
    
    If Index < mCount - 1 Then
        Dim SourceIndex     As Long
        Dim NumberToCopy    As Long
        SourceIndex = Index + 1
        NumberToCopy = (mCount - Index) * vbSizeOfVariant
        
        CopyMemory mKeys(Index), mKeys(SourceIndex), NumberToCopy
        CopyMemory mValues(Index), mValues(SourceIndex), NumberToCopy
        
        Dim LastIndex As Long
        LastIndex = mCount - 1
        ZeroMemory mKeys(LastIndex), vbSizeOfVariant
        ZeroMemory mValues(LastIndex), vbSizeOfVariant
    End If
    
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

''
' Replaces the value at a specific index in a <a href="SortedList.html">SortedList</a> object.
'
' @param Index The zero-based index at which to set the value.
' @param Value The value to save into the <a href="SortedList.html">SortedList</a> object.
' @remarks The index sequence is based on the sort sequence. When an element is added, it is inserted into <a href="SortedList.html">SortedList</a> in the correct sort order,
' and the indexing adjusts accordingly. When an element is removed, the indexing also adjusts accordingly. Therefore, the index of a specific
' key/value pair might change as elements are added or removed from the <a href="SortedList.html">SortedList</a> object.
' @include "..\DocIncludes\SortedList.SetByIndex.html"
Public Sub SetByIndex(ByVal Index As Long, ByRef Value As Variant)
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    VariantCopyInd mValues(Index), Value
    mVersion = mVersion + 1
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Sets the capacity to the actual number of elements in a <a href="SortedList.html">SortedList</a> object.
'
' @remarks This method can be used to minimize a collection's memory overhead if no new elements will be added to the collection.
' <p>To reset a <a href="SortedList.html">SortedList</a> object to its initial state, call the <a href="SortedList.Clear.html">Clear</a> method
' before calling <b>TrimToSize</b>. Trimming an empty <a href="SortedList.html">SortedList</a> sets the capacity of the <a href="SortedList.html">SortedList</a> to the default capacity.</p>
' @see #Clear
' @see #Capacity
' @see #Count
' @include "..\DocIncludes\SortedList.TrimToSize.html"
Public Sub TrimToSize()
    Capacity = mCount
End Sub

''
' Gets the values in the <a href="SortedList.html">SortedList</a> object.
'
' @return An ICollection object containing the values in the <a href="SortedList.html">SortedList</a> object.
' @remarks The <a href="ICollection.html">ICollection</a> object is a read-only view of the values of the <a href="SortedList.html">SortedList</a> object.
' Modifications made to the underlying <a href="SortedList.html">SortedList</a> are immediately reflected in the <a href="ICollection.html">ICollection</a>.
' <p>The elements of the <a href="ICollection.html">ICollection</a> are sorted in the same order as the values of the <a href="SortedList.html">SortedList</a>.</p>
' <p>This property is similar to the <a href="SortedKey.GetValueList.html">GetValueList</a> method,
' but returns an <a href="ICollection.html">ICollection</a> object instead of an <a href="IList.html">IList</a> object.</p>
' @see ICollection
' @see #Keys
' @see #GetValueList
Public Property Get Values() As ICollection
    Set Values = NewSortedValueList
End Property

Friend Function GetKeyEnumerator() As IEnumerator
    Dim Ret As New SortedListEnumerator
    Ret.Init Me, mKeys, mValues, slKeys
    Set GetKeyEnumerator = Ret
End Function

Friend Function GetValueEnumerator() As IEnumerator
    Dim Ret As New SortedListEnumerator
    Ret.Init Me, mKeys, mValues, slValues
    Set GetValueEnumerator = Ret
End Function

Friend Sub CopyKeysTo(ByRef Arr As Variant, ByVal Index As Long)
    CorArray.CopyEx mKeys, 0, Arr, Index, mCount
End Sub

Friend Sub CopyValuesTo(ByRef Arr As Variant, ByVal Index As Long)
    CorArray.CopyEx mValues, 0, Arr, Index, mCount
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As IDictionary, ByVal Comparer As IComparer, ByVal Capacity As Long)
    Set mComparer = Comparer
    
    If List Is Nothing Then
        Me.Capacity = Capacity
    Else
        mCount = List.Count
        Me.Capacity = mCount
        List.Keys.CopyTo mKeys, 0
        List.Values.CopyTo mValues, 0
        CorArray.SortKeyEx mKeys, mValues, 0, mCount, Comparer
    End If
End Sub

Friend Sub InitClone(ByRef Keys() As Variant, ByRef Values() As Variant, ByVal Count As Long, ByVal Comparer As IComparer)
    mCount = Count
    Set mComparer = Comparer
    mKeys = Keys
    mValues = Values
    mCapacity = UBound(Keys) + 1
End Sub

Private Sub Class_InitProperties()
    mKeys = Array()
    mValues = Array()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mCount = .ReadProperty(PropCount, 0)
        EnsureCapacity .ReadProperty(PropCapacity, DefaultCapacity)
        Set mComparer = .ReadProperty(PropComparer, Nothing)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Helper.MoveVariant mKeys(i), .ReadProperty(MakeKeyKey(i))
            Helper.MoveVariant mValues(i), .ReadProperty(MakeValueKey(i))
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropCount, mCount
        .WriteProperty PropCapacity, mCapacity
        .WriteProperty PropComparer, mComparer
        
        Dim i As Long
        For i = 0 To mCount - 1
            .WriteProperty MakeKeyKey(i), mKeys(i)
            .WriteProperty MakeValueKey(i), mValues(i)
        Next i
    End With
End Sub

Private Function MakeKeyKey(ByVal Index As Long) As String
    MakeKeyKey = "Key" & Index
End Function

Private Function MakeValueKey(ByVal Index As Long) As String
    MakeValueKey = "Value" & Index
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Insert(ByVal Index As Long, ByRef Key As Variant, ByRef Value As Variant)
    EnsureCapacity mCount + 1
    
    If Index < mCount Then
        Dim DestIndex       As Long
        Dim NumberToCopy    As Long
        DestIndex = Index + 1
        NumberToCopy = (mCount - Index) * vbSizeOfVariant
    
        CopyMemory mKeys(DestIndex), mKeys(Index), NumberToCopy
        CopyMemory mValues(DestIndex), mValues(Index), NumberToCopy
        
        ZeroMemory mKeys(Index), vbSizeOfVariant
        ZeroMemory mValues(Index), vbSizeOfVariant
    End If
    
    VariantCopyInd mKeys(Index), Key
    VariantCopyInd mValues(Index), Value
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    
    If mCapacity = 0 Then
        NewCapacity = DefaultCapacity
    Else
        NewCapacity = mCapacity * 2
    End If
    
    If RequiredCapacity > NewCapacity Then
        NewCapacity = RequiredCapacity
    End If
    
    ReDim Preserve mKeys(0 To NewCapacity - 1)
    ReDim Preserve mValues(0 To NewCapacity - 1)
    mCapacity = NewCapacity
End Sub

Private Function NewSortedKeyList() As SortedKeyList
    Set NewSortedKeyList = New SortedKeyList
    NewSortedKeyList.Init Me
End Function

Private Function NewSortedValueList() As SortedValueList
    Set NewSortedValueList = New SortedValueList
    NewSortedValueList.Init Me
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
'   IDictionary
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IDictionary_Add(Key As Variant, Value As Variant)
    Add Key, Value
End Sub

Private Sub IDictionary_Clear()
    Clear
End Sub

Private Function IDictionary_Contains(Key As Variant) As Boolean
    IDictionary_Contains = Contains(Key)
End Function

Private Sub IDictionary_CopyTo(Arr As Variant, ByVal Index As Long)
    CopyTo Arr, Index
End Sub

Private Property Get IDictionary_Count() As Long
    IDictionary_Count = Count
End Property

Private Function IDictionary_GetEnumerator() As IDictionaryEnumerator
    Set IDictionary_GetEnumerator = GetEnumerator
End Function

Private Property Get IDictionary_IsFixedSize() As Boolean
    IDictionary_IsFixedSize = IsFixedSize
End Property

Private Property Get IDictionary_IsReadOnly() As Boolean
    IDictionary_IsReadOnly = IsReadOnly
End Property

Private Property Set IDictionary_Item(Key As Variant, Value As Variant)
    Set Item(Key) = Value
End Property

Private Property Let IDictionary_Item(Key As Variant, Value As Variant)
    Item(Key) = Value
End Property

Private Property Get IDictionary_Item(Key As Variant) As Variant
    Helper.MoveVariant IDictionary_Item, Item(Key)
End Property

Private Property Get IDictionary_Keys() As ICollection
    Set IDictionary_Keys = Keys
End Property

Private Function IDictionary_NewEnum() As stdole.IUnknown
    Set IDictionary_NewEnum = NewEnum
End Function

Private Sub IDictionary_Remove(Key As Variant)
    Remove Key
End Sub

Private Property Get IDictionary_Values() As ICollection
    Set IDictionary_Values = Values
End Property


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
