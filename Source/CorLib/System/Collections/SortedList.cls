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
'    Module: SortedList
'

''
' Provides a means to maintain a list of sorted items based on an associated key.
'
' @remarks The SortedList maintains sort order for every key/value pair
' added to the list. The list is expanded as needed with the addition of new items.
' If a key already exists in the list, an exception is thrown.
'
' @see Constructors
' @see IDictionary
' @see ICollection
' @see IEnumerable
'
Option Explicit
Implements IObject
Implements IEnumerable
Implements ICollection
Implements IDictionary

Private Const DEF_CAPACITY As Long = 16

Private mKeys()     As Variant
Private mValues()   As Variant
Private mComparer   As IComparer
Private mCapacity   As Long
Private mCount      As Long
Private mVersion    As Long


''
' Adds a key/value pair to the list.
'
' @param key The value to used to maintain a sorted order.
' @param value The value associated with the key.
' @remark key must be unique.
Public Sub Add(ByRef Key As Variant, ByRef Value As Variant)
    Dim i As Long
    i = cArray.BinarySearchEx(mKeys, Key, 0, mCount, mComparer)
    If i >= 0 Then _
        Throw Cor.NewArgumentException("An element with this key already exists.", "key")
    
    i = Not i
    Call EnsureCapacity(mCount + 1)
    If i < mCount Then
        Call CopyMemory(mKeys(i + 1), mKeys(i), (mCount - i) * 16)
        Call ZeroMemory(mKeys(i), 16)
        Call CopyMemory(mValues(i + 1), mValues(i), (mCount - i) * 16)
        Call ZeroMemory(mValues(i), 16)
    End If
    Call VariantCopyInd(mKeys(i), Key)
    Call VariantCopyInd(mValues(i), Value)
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

''
' Returns the size of the underlying array.
'
' @return The number of elements the underlying array can hold
' before needing to be resized.
'
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the size of the underlying array.
'
' @param RHS The number of elements the underlying array can contain
' before it needs to be resized.
' @remarks If the capacity is set less than the number of items in the
' array, an ArgumentOutOfRangeException exception will be thrown.
'
Public Property Let Capacity(ByVal RHS As Long)
    If RHS < mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", Message:=Environment.GetResourceString(ArgumentOutOfRange_SmallCapacity))
    If RHS = mCapacity Then Exit Property
    If RHS = 0 Then RHS = DEF_CAPACITY
    mCapacity = RHS
    ReDim Preserve mKeys(0 To mCapacity - 1)
    ReDim Preserve mValues(0 To mCapacity - 1)
    mVersion = mVersion + 1
End Property

''
' Removes all of the items from the list.
'
' @remarks The Count is set to 0, but the capacity remains unchanged.
'
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
'
Public Function Clone() As SortedList
    Set Clone = New SortedList
    Call Clone.CloneHelper(mKeys, mValues, mCount, mComparer)
End Function

''
' Searches the list to determine if the key is contained in the list.
'
' @param Key The key to check if exists in the list.
' @return Indication of the key existing in the list.
' @remarks Internally this function calls <b>ContainsKey</b>.
'
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = ContainsKey(Key)
End Function

''
' Searches the list to determine if the key is contained in the list.
'
' @param Key The key to check if exists in the list.
' @return Indication of the key existing in the list.
'
Public Function ContainsKey(ByRef Key As Variant) As Boolean
    ContainsKey = (cArray.BinarySearchEx(mKeys, Key, 0, mCount, mComparer) >= 0)
End Function

''
' Searches the list to determine if the value is contained in the list.
'
' @param value The value to search for in the list.
' @param comparer A user-supplied IComparer object to perform special comparisons as necessary.
' @return Indication of the value being found in the list.
'
Public Function ContainsValue(ByRef Value As Variant) As Boolean
    ContainsValue = cArray.IndexOf(mValues, Value, 0, mCount) >= 0
End Function

''
' Creates a DictionaryEntry for each Key/Value pair and places it in the array.
'
' @param dstArray The array to copy the DictionaryEntry objects to.
' @param Index The starting index in dstArray to begin copying to.
' @remarks Since DictionaryEntry objects are being stored in the array,
' the array must be able to accept that type of object.
'
Public Sub CopyTo(ByRef DstArray As Variant, ByVal Index As Long)
    Dim i As Long
    
    If cArray.IsNull(DstArray) Then _
        Throw Cor.NewArgumentNullException("dstArray", Environment.GetResourceString(ArgumentNull_Array))
    If Index < LBound(DstArray) Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_LBound))
    If cArray.GetRank(DstArray) <> 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimNotSupported))
    If Index + mCount - 1 > UBound(DstArray) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ArrayPlusOffTooSmall))
    
    For i = 0 To mCount - 1
        Set DstArray(Index + i) = Cor.NewDictionaryEntry(mKeys(i), mValues(i))
    Next i
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
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Retrieves a value in the list by an index.
'
' @param Index The index into the list from which to retrieve the value.
' @return The value at the specified index in the list.
' @remarks Since the list maintains sort order with each addition
' and removal of a key/value pair, there is no garauntee the same value
' will be returned by the same index in future calls.
'
Public Function GetByIndex(ByVal Index As Long) As Variant
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    Call VariantCopy(GetByIndex, mValues(Index))
End Function

''
' Returns an enumerator that can be used to iterate through the key/value pairs.
'
' @return An IEnumerator object to be use for iterating the key/value pairs.
'
Public Function GetEnumerator() As IEnumerator
    Dim Ret As New SortedListEnumerator
    Call Ret.Init(Me, SAPtr(mKeys), SAPtr(mValues), slEntries)
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
'
Public Function GetKey(ByVal Index As Long) As Variant
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    Call VariantCopy(GetKey, mKeys(Index))
End Function

''
' Returns a ReadOnly IList object that is used to access the keys in the list.
'
' @return An IList object containing the keys in the list.
' @remarks The IList object maintains a reference to the original sorted list.
' It will reflect any changes in the sorted list.
'
Public Function GetKeyList() As IList
    Set GetKeyList = NewSortedKeyList
End Function

''
' Returns a ReadOnly IList object that is used to access the values in the list.
'
' @return An IList object used to access the values in the list.
' @remarks The IList object maintains a reference to the original sorted list.
' Any changes in the sorted list will be reflected in the IList object.
'
Public Function GetValueList() As IList
    Set GetValueList = NewSortedValueList
End Function

''
' Searches for a key in the list and returns the index it was found.
'
' @param Key The key to find the index of.
' @return The index for the specified key. If a -1 is returned, then
' the key was not found in the list.
'
Public Function IndexOfKey(ByRef Key As Variant) As Long
    IndexOfKey = cArray.IndexOf(mKeys, Key, 0, mCount)
End Function

''
' Searches for a value in the list and returns the index it was found.
'
' @param Value The value to search for.
' @param Comparer A user-supplied comparer for special comparison, especially for user-defined types.
' @return The index the value was found, or a -1 if the value was not found.
'
Public Function IndexOfValue(ByRef Value As Variant) As Long
    IndexOfValue = cArray.IndexOf(mValues, Value, 0, mCount)
End Function

''
' Returns if the list can change in size or not.
'
' @return Indication if the list is fixed or not.
' @remarks Fixed size means items cannot be added or removed from the list.
' It does not prevent the existing items from being modified.
'
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Returns if the list can be modified in any way.
'
' @return Indication if the list can be modified.
' @remarks A ReadOnly list cannot have items added or removed, nor
' can the existing items be altered. This does not mean properties
' on objects in the list are prohibited.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns the value associated with the given key.
'
' @param Key The key of the value to be returned.
' @return The value associated with the key.
' @remarks If the key was not found, then Empty is returned.
'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = cArray.BinarySearchEx(mKeys, Key, 0, mCount, mComparer)
    If i < 0 Then Exit Property
    Call VariantCopy(Item, mValues(i))
End Property

''
' Sets a value to a key.
'
' @param Key The key to associate the value with.
' @param Value The value to be associated to the key.
' @remarks If the key already exists in the list, then that value
' is replaced with the new value. If the key does not exist, then
' it is added with the associated value.
'
Public Property Let Item(ByRef Key As Variant, ByRef Value As Variant)
    Dim i As Long
    i = cArray.BinarySearchEx(mKeys, Key, 0, mCount, mComparer)
    If i < 0 Then
        Call Add(Key, Value)
    Else
        mValues(i) = Value
        mVersion = mVersion + 1
    End If
End Property

''
' Sets a value to a key.
'
' @param Key The key to associate the value with.
' @param Value The value to be associated to the key.
' @remarks If the key already exists in the list, then that value
' is replaced with the new value. If the key does not exist, then
' it is added with the associated value.
'
Public Property Set Item(ByRef Key As Variant, ByRef Value As Variant)
    Dim i As Long
    i = cArray.BinarySearchEx(mKeys, Key, 0, mCount, mComparer)
    If i < 0 Then
        Call Add(Key, Value)
    Else
        Set mValues(i) = Value
        mVersion = mVersion + 1
    End If
End Property

''
' Returns the keys in the sorted list.
'
' @return An ICollection object containing the keys.
'
Public Property Get Keys() As ICollection
    Set Keys = NewSortedKeyList
End Property

''
' Returns a For..Each compatible enumerator.
'
' @return The enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Removes a key/value pair based on the key.
'
' @param Key The key used to find the value to be removed.
'
Public Sub Remove(ByRef Key As Variant)
    Dim i As Long
    If mCount = 0 Then Exit Sub
    i = cArray.BinarySearchEx(mKeys, Key, 0, mCount, mComparer)
    If i >= 0 Then Call RemoveAt(i)
End Sub

''
' Removes a key/value pair at a specific index in the list.
'
' @param Index The index of the key/value pair to be removed.
'
Public Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    mKeys(Index) = Empty
    mValues(Index) = Empty
    If Index < mCount - 1 Then
        Call CopyMemory(mKeys(Index), mKeys(Index + 1), (mCount - Index) * 16)
        Call ZeroMemory(mKeys(mCount - 1), 16)
        Call CopyMemory(mValues(Index), mValues(Index + 1), (mCount - Index) * 16)
        Call ZeroMemory(mValues(mCount - 1), 16)
    End If
    mCount = mCount - 1
    mVersion = mVersion + 1
End Sub

''
' Sets the value at a specific index in the list.
'
' @param Index The index at which to set the value.
' @param Value The value to set at the specified index.
' @remarks The key at the index will be associated with the new value.
'
Public Sub SetByIndex(ByVal Index As Long, ByRef Value As Variant)
    If Index < 0 Or Index >= mCount Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Index, Environment.GetResourceString(ArgumentOutOfRange_Index))
    Call VariantCopyInd(mValues(Index), Value)
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
' Sets the capacity to the number of items in the list.
'
' @remarks This is used to release extra space being used by
' the internal array. The capacity is set to the number of
' items already in the list.
'
Public Sub TrimToSize()
    Capacity = mCount
End Sub

''
' Returns a list of the values in the list.
'
' @return An ICollection object used to access the values.
'
Public Property Get Values() As ICollection
    Set Values = NewSortedValueList
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetKeyEnumerator() As IEnumerator
    Dim Ret As New SortedListEnumerator
    Call Ret.Init(Me, SAPtr(mKeys), SAPtr(mValues), slKeys)
    Set GetKeyEnumerator = Ret
End Function

Friend Function GetValueEnumerator() As IEnumerator
    Dim Ret As New SortedListEnumerator
    Call Ret.Init(Me, SAPtr(mKeys), SAPtr(mValues), slValues)
    Set GetValueEnumerator = Ret
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    ReDim mKeys(0 To DEF_CAPACITY - 1)
    ReDim mValues(0 To DEF_CAPACITY - 1)
    mCapacity = DEF_CAPACITY
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mCount = .ReadProperty("Count", 0)
        Call EnsureCapacity(.ReadProperty("Capacity", DEF_CAPACITY))
        Set mComparer = .ReadProperty("Comparer", Nothing)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Call Helper.MoveVariant(mKeys(i), .ReadProperty("Key" & i))
            Call Helper.MoveVariant(mValues(i), .ReadProperty("Value" & i))
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Count", mCount)
        Call .WriteProperty("Capacity", mCapacity)
        Call .WriteProperty("Comparer", mComparer)
        
        Dim i As Long
        For i = 0 To mCount - 1
            Call .WriteProperty("Key" & i, mKeys(i))
            Call .WriteProperty("Value" & i, mValues(i))
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

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    If RequiredCapacity <= mCapacity Then Exit Sub
    NewCapacity = mCapacity * 2
    If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
    ReDim Preserve mKeys(0 To NewCapacity - 1)
    ReDim Preserve mValues(0 To NewCapacity - 1)
    mCapacity = NewCapacity
End Sub

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
'   IDictionary Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IDictionary_Add(Key As Variant, Value As Variant)
    Call Add(Key, Value)
End Sub

Private Sub IDictionary_Clear()
    Call Clear
End Sub

Private Function IDictionary_Contains(Key As Variant) As Boolean
    IDictionary_Contains = Contains(Key)
End Function

Private Sub IDictionary_CopyTo(Arr As Variant, ByVal Index As Long)
    Call CopyTo(Arr, Index)
End Sub

Private Property Get IDictionary_Count() As Long
    IDictionary_Count = Count
End Property

Private Function IDictionary_GetEnumerator() As IEnumerator
    Set IDictionary_GetEnumerator = GetEnumerator
End Function

Private Property Get IDictionary_IsFixedSize() As Boolean
    IDictionary_IsFixedSize = IsFixedSize
End Property

Private Property Get IDictionary_IsReadOnly() As Boolean
    IDictionary_IsReadOnly = IsReadOnly
End Property

Private Property Set IDictionary_Item(Key As Variant, RHS As Variant)
    Set Item(Key) = RHS
End Property

Private Property Let IDictionary_Item(Key As Variant, RHS As Variant)
    Item(Key) = RHS
End Property

Private Property Get IDictionary_Item(Key As Variant) As Variant
    Call Helper.MoveVariant(IDictionary_Item, Item(Key))
End Property

Private Property Get IDictionary_Keys() As ICollection
    Set IDictionary_Keys = Keys
End Property

Private Function IDictionary_NewEnum() As stdole.IUnknown
    Set IDictionary_NewEnum = NewEnum
End Function

Private Sub IDictionary_Remove(Key As Variant)
    Call Remove(Key)
End Sub

Private Property Get IDictionary_Values() As ICollection
    Set IDictionary_Values = Values
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function

Private Function NewSortedKeyList() As SortedKeyList
    Set NewSortedKeyList = New SortedKeyList
    Call NewSortedKeyList.Init(Me)
End Function

Private Function NewSortedValueList() As SortedValueList
    Set NewSortedValueList = New SortedValueList
    Call NewSortedValueList.Init(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub CloneHelper(ByRef Keys() As Variant, ByRef Values() As Variant, ByVal Count As Long, ByVal Comparer As IComparer)
    mCount = Count
    Set mComparer = Comparer
    mKeys = Keys
    mValues = Values
    mCapacity = UBound(Keys) + 1
End Sub

Friend Sub CopyKeysTo(ByRef DstArray As Variant, ByVal Index As Long)
    Call cArray.CopyEx(mKeys, 0, DstArray, Index, mCount)
End Sub

Friend Sub CopyValuesTo(ByRef DstArray As Variant, ByVal Index As Long)
    Call cArray.CopyEx(mValues, 0, DstArray, Index, mCount)
End Sub

Friend Sub Init(ByVal List As IDictionary, ByVal Comparer As IComparer, ByVal Capacity As Long)
    Set mComparer = Comparer
    If List Is Nothing Then
        Me.Capacity = Capacity
    Else
        mCount = List.Count
        Me.Capacity = mCount
        Call List.Keys.CopyTo(mKeys, 0)
        Call List.Values.CopyTo(mValues, 0)
        Call cArray.SortKeyEx(mKeys, mValues, 0, mCount, Comparer)
    End If
End Sub

Friend Property Get Version() As Long
    Version = mVersion
End Property
