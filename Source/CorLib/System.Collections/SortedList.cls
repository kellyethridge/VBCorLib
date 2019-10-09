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
' Adds a key/value pair to the list.
'
' @param key The value to used to maintain a sorted order.
' @param value The value associated with the key.
' @remark key must be unique.
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
' @param Value The number of elements the underlying array can contain
' before it needs to be resized.
' @remarks If the capacity is set less than the number of items in the
' array, an ArgumentOutOfRangeException exception will be thrown.
'
Public Property Let Capacity(ByVal Value As Long)
    If Value < mCount Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_SmallCapacity
    
    If Value > 0 Then
        ReDim Preserve mKeys(0 To Value - 1)
        ReDim Preserve mValues(0 To Value - 1)
    Else
        mKeys = Array()
        mValues = Array()
    End If
    
    mCapacity = Value
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
    Clone.InitClone mKeys, mValues, mCount, mComparer
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
    ContainsKey = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer) >= 0
End Function

''
' Searches the list to determine if the value is contained in the list.
'
' @param value The value to search for in the list.
' @param comparer A user-supplied IComparer object to perform special comparisons as necessary.
' @return Indication of the value being found in the list.
'
Public Function ContainsValue(ByRef Value As Variant) As Boolean
    ContainsValue = CorArray.IndexOf(mValues, Value, 0, mCount) >= 0
End Function

''
' Creates a DictionaryEntry for each Key/Value pair and places it in the array.
'
' @param Arr The array to copy the DictionaryEntry objects to.
' @param ArrayIndex The starting index in dstArray to begin copying to.
' @remarks Because DictionaryEntry objects are being stored in the array,
' the array must be able to accept that type of object.
'
Public Sub CopyTo(ByRef Arr As Variant, ByVal ArrayIndex As Long)
    If CorArray.IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    If CorArray.Rank(Arr) <> 1 Then _
        Error.Argument Arg_RankMultiDimNotSupported, "Arr"
    If ArrayIndex < LBound(Arr) Then _
        Error.ArgumentOutOfRange "ArrayIndex", ArgumentOutOfRange_LBound
    If ArrayIndex + mCount - 1 > UBound(Arr) Then _
        Error.Argument Arg_ArrayPlusOffTooSmall
    
    Dim i As Long
    For i = 0 To mCount - 1
        Set Arr(ArrayIndex + i) = Cor.NewDictionaryEntry(mKeys(i), mValues(i))
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
'
Public Function GetKey(ByVal Index As Long) As Variant
    If Index < 0 Or Index >= mCount Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
        
    VariantCopy GetKey, mKeys(Index)
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
    Dim Index As Long
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index < 0 Then
        Index = -1
    End If
    
    IndexOfKey = Index
End Function

''
' Searches for a value in the list and returns the index it was found.
'
' @param Value The value to search for.
' @param Comparer A user-supplied comparer for special comparison, especially for user-defined types.
' @return The index the value was found, or a -1 if the value was not found.
'
Public Function IndexOfValue(ByRef Value As Variant) As Long
    IndexOfValue = CorArray.IndexOf(mValues, Value, 0, mCount)
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
    Dim Index As Long
    Index = CorArray.InternalBinarySearch(mKeys, 0, mCount, Key, mComparer)
    
    If Index >= 0 Then
        VariantCopy Item, mValues(Index)
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
' Sets a value to a key.
'
' @param Key The key to associate the value with.
' @param Value The value to be associated to the key.
' @remarks If the key already exists in the list, then that value
' is replaced with the new value. If the key does not exist, then
' it is added with the associated value.
'
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
' Removes a key/value pair at a specific index in the list.
'
' @param Index The index of the key/value pair to be removed.
'
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
' Sets the value at a specific index in the list.
'
' @param Index The index at which to set the value.
' @param Value The value to set at the specified index.
' @remarks The key at the index will be associated with the new value.
'
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
