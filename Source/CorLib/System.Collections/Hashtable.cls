VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hashtable"
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
' Module: Hashtable
'

''
' Represents a collection of key/value pairs that are organized based on the hash code of the key.
'
' @see Constructors
' @see IDictionary
' @see ICollection
' @see IEnumerable
' @see ICloneable
' @see CaseInsensitiveComparer
' @include "..\DocIncludes\Hashtable.html"
Option Explicit
Implements IObject
Implements IDictionary
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IVersionable

Private Const PropCount                 As String = "Count"
Private Const PropKey                   As String = "Key"
Private Const PropValueSubtype          As String = "ValueSubType"
Private Const PropValue                 As String = "Value"
Private Const PropComparer              As String = "Comparer"
Private Const PropUseDefaultComparer    As String = "UseDefaultComparer"
Private Const SubtypeNormal             As Long = 0
Private Const SubtypeEmpty              As Long = 1
Private Const SubtypeNothing            As Long = 2
Private Const DefaultLoadFactor         As Single = 1!

Private mLoadFactor     As Single
Private mLoadSize       As Long
Private mCount          As Long
Private mBuckets()      As HashBucket
Private mCapacity       As Long
Private mVersion        As Long
Private mKeyComparer    As IEqualityComparer


''
' Adds an element with the specified key and value into the <a href="Hashtable.html">Hashtable</a>.
'
' @param Key The key of the element to add.
' @param Value The value of the element to add.
' @see #Remove
' @see #Item
' @include "..\DocIncludes\Hashtable.Add.html"
Public Sub Add(ByRef Key As Variant, ByRef Value As Variant)
    InsertItem Key, Value, True, False
End Sub

''
' Removes all elements from the <a href="Hashtable.html">Hashtable</a>.
'
' @include "..\DocIncludes\Hashtable.Clear.html"
Public Sub Clear()
    ReDim mBuckets(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Removes <b>Deleted</b> buckets from the internal array.
'
' @remarks When many items have been deleted from the table, the
' buckets that contained the items are marked as Deleted. When many
' buckets are marked as <b>Deleted</b>, the search routine cannot stop
' searching the buckets for a duplicate key even if an empty Deleted
' bucket has been found. So to prevent the unnessecary searching past
' Deleted buckets, the table should be cleaned to remove the Deleted
' buckets and only have <b>Empty</b> and <i>Occupied</b> buckets remaining.
' <p>This is recommended when a large number of items are deleted.</p>
Public Sub Clean()
    ReinsertItems
    mVersion = mVersion + 1
End Sub

''
' Creates a shallow copy of the <a href="Hashtable.html">Hashtable</a>.
'
' @return A shallow copy of the <a href="Hashtable.html">Hashtable</a>.
' @remarks A shallow copy of a collection copies only the elements of the collection, whether they are reference types or value types,
' but it does not copy the objects that the references refer to. The references in the new collection point to the same objects that
' the references in the original collection point to.
' <p>In contrast, a deep copy of a collection copies the elements and everything directly or indirectly referenced by the elements.</p>
' <p>The Hashtable clone has the same count, the same capacity, the same <a href="IEqualityComparer.html">IEqualityComparer</a>
' implementation, and the same IComparer implementation as the original <a href="Hashtable.html">Hashtable</a>.</p>
' <p>This method is an O(<i>n</i>) operation, where <i>n</i> is <a href="Hashtable.Get.Count.html">Count</a>.</p>
' @see #CopyTo
Public Function Clone() As Hashtable
    Set Clone = New Hashtable
    Clone.InitClone mBuckets, mCount, mLoadFactor, mLoadSize, mKeyComparer
End Function

''
' Determines whether the <a href="Hashtable.html">Hashtable</a> contains a specific key.
'
' @param Key The key to locate in the <a href="Hashtable.html">Hashtable</a>.
' @return <b>True</b> if the <a href="Hashtable.html">Hashtable</a> contains an element with the specified key; otherwise, <b>False</b>.
' @see #ContainsKey
' @include "..\DocIncludes\Hashtable.Contains.html"
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = ContainsKey(Key)
End Function

''
' Determines whether the <a href="Hashtable.html">Hashtable</a> contains a specific key.
'
' @param Key The key to locate in the <a href="Hashtable.html">Hashtable</a>.
' @return <b>True</b> if the <a href="Hashtable.html">Hashtable</a> contains an element with the specified key; otherwise, <b>False</b>.
' @see #Contains
' @see #ContainsValue
' @include "..\DocIncludes\Hashtable.ContainsKey.html"
Public Function ContainsKey(ByRef Key As Variant) As Boolean
    ContainsKey = (GetKeyIndex(Key) >= 0)
End Function

''
' Determines whether the <a href="Hashtable.html">Hashtable</a> contains a specific value.
'
' @param Value The value to locate in the <a href="Hashtable.html">Hashtable</a>.
' @return <b>True</b> if the <a href="Hashtable.html">Hashtable</a> contains an element with the specified value; otherwise, <b>False</b>.
' @see #ContainsKey
' @include "..\DocIncludes\Hashtable.ContainsValue.html"
Public Function ContainsValue(ByRef Value As Variant) As Boolean
    Dim i As Long
    
    For i = 0 To mCapacity - 1
        With mBuckets(i)
            If .State = OccupiedState Then
                If Object.Equals(.Value, Value) Then
                    ContainsValue = True
                    Exit Function
                End If
            End If
        End With
    Next i
End Function

''
' Copies the <a href="Hashtable.html">Hashtable</a> elements to a one-dimensional Array instance at the specified index.
'
' @param Arr The one-dimensional array that is the destination of the <a href="DictionaryEntry.html">DictionaryEntry</a>
' objects copied from <a href="Hashtable.html">Hashtable</a>. The Array must have zero-based indexing.
' @param ArrayIndex The lowerbound-based index in <i>Arr</i> at which copying begins.
' @see DictionaryEntry
' @include "..\DocIncludes\Hashtable.CopyTo.html"
Public Sub CopyTo(ByRef Arr As Variant, ByVal ArrayIndex As Long)
    CopyToCore Arr, ArrayIndex, EnumEntries
End Sub

''
' Gets the number of key/value pairs contained in the <a href="Hashtable.html">Hashtable</a>.
'
' @return The number of key/value pairs contained in the <a href="Hashtable.html">Hashtable</a>.
' @remarks Retrieving the value of this property is an O(1) operation.
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
' Returns an <a href="IDictionaryEnumerator.html">IDictionaryEnumerator</a> that iterates through the <a href="Hashtable.html">Hashtable</a>.
'
' @return An <a href="IDictionaryEnumerator.html">IDictionaryEnumerator</a> for the <a href="Hashtable.html">Hashtable</a>.
' @remarks The elements enumerated are <a href="DictionaryEntry.html">DictionaryEntry</a> objects each containing a key/value pair.
Public Function GetEnumerator() As IDictionaryEnumerator
    Dim Ret As New HashtableEnumerator
    Ret.Init Me, mBuckets, EnumEntries
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
' Gets a value indicating whether the <a href="Hashtable.html">Hashtable</a> has a fixed size.
'
' @return <b>True</b> if the <a href="Hashtable.html">Hashtable</a> is read-only; otherwise, <b>False</b>. The default is <b>False</b>.
' @remarks This implements the <a href="IList.Get.IsFixedSize.html">IList.IsFixedSize</a>.
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Gets a value indicating whether the <a href="Hashtable.html">Hashtable</a> is read-only.
'
' @return <b>True</b> if the <a href="Hashtable.html">Hashtable</a> is read-only; otherwise, <b>False</b>. The default is <b>False</b>.
' @remarks This implements the <a href="IList.Get.IsReadOnly.html">IList.IsReadOnly</a>.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Gets the value associated with the specified key.
'
' @param Key The key whose value to get.
' @return The value associated with the specified key. If the specified key is not found, attempting to get it returns <b>Empty</b>, and attempting to set it creates a new element using the specified key.
' @remarks This property provides the ability to access a specific element in the collection by using the following syntax: <code>myCollection(key)</code>.
' <p>You can also use the <b>Item</b> property to add new elements by setting the value of a key that does not exist in the <a href="Hashtable.html">Hashtable</a>; for example,
' <code>myCollection("myNonexistentKey") = myValue</code>. However, if the specified key already exists in the <a href="Hashtable.html">Hashtable</a>,
' setting the <b>Item</b> property overwrites the old value. In contrast, the <a href="Hashtable.Add.html">Add</a> method does not modify existing elements.</p>
' <p>To distinguish between <b>Empty</b> that is returned because the specified key is not found and <b>Empty</b> that is returned because the value of the specified key is <b>Empty</b>,
' use the <a href="Hashtable.Contains.html">Contains</a> method or the <a href="Hashtable.ContainsKey.html">ContainsKey</a> method to determine if the key exists in the list.</p>
' <p>Retrieving the value of this property is an O(1) operation; setting the property is also an O(1) operation.</p>
' @see #Add
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Index As Long
    Index = GetKeyIndex(Key)
    If Index >= 0 Then
        VariantCopy Item, mBuckets(Index).Value
    End If
End Property

''
' Sets the value associated with the specified key.
'
' @param Key The key whose value to get.
' @param Value The value associated with the specified key. If the specified key is not found, attempting to get it returns <b>Empty</b>, and attempting to set it creates a new element using the specified key.
' @remarks This property provides the ability to access a specific element in the collection by using the following syntax: <code>myCollection(key)</code>.
' <p>You can also use the <b>Item</b> property to add new elements by setting the value of a key that does not exist in the <a href="Hashtable.html">Hashtable</a>; for example,
' <code>myCollection("myNonexistentKey") = myValue</code>. However, if the specified key already exists in the <a href="Hashtable.html">Hashtable</a>,
' setting the <b>Item</b> property overwrites the old value. In contrast, the <a href="Hashtable.Add.html">Add</a> method does not modify existing elements.</p>
' <p>To distinguish between <b>Empty</b> that is returned because the specified key is not found and <b>Empty</b> that is returned because the value of the specified key is <b>Empty</b>,
' use the <a href="Hashtable.Contains.html">Contains</a> method or the <a href="Hashtable.ContainsKey.html">ContainsKey</a> method to determine if the key exists in the list.</p>
' <p>Retrieving the value of this property is an O(1) operation; setting the property is also an O(1) operation.</p>
' @see #Add
Public Property Let Item(ByRef Key As Variant, ByRef Value As Variant)
    InsertItem Key, Value, False, False
End Property

''
' Sets the value associated with the specified key.
'
' @param Key The key whose value to get.
' @param Value The value associated with the specified key. If the specified key is not found, attempting to get it returns <b>Empty</b>, and attempting to set it creates a new element using the specified key.
' @remarks This property provides the ability to access a specific element in the collection by using the following syntax: <code>myCollection(key)</code>.
' <p>You can also use the <b>Item</b> property to add new elements by setting the value of a key that does not exist in the <a href="Hashtable.html">Hashtable</a>; for example,
' <code>myCollection("myNonexistentKey") = myValue</code>. However, if the specified key already exists in the <a href="Hashtable.html">Hashtable</a>,
' setting the <b>Item</b> property overwrites the old value. In contrast, the <a href="Hashtable.Add.html">Add</a> method does not modify existing elements.</p>
' <p>To distinguish between <b>Empty</b> that is returned because the specified key is not found and <b>Empty</b> that is returned because the value of the specified key is <b>Empty</b>,
' use the <a href="Hashtable.Contains.html">Contains</a> method or the <a href="Hashtable.ContainsKey.html">ContainsKey</a> method to determine if the key exists in the list.</p>
' <p>Retrieving the value of this property is an O(1) operation; setting the property is also an O(1) operation.</p>
' @see #Add
Public Property Set Item(ByRef Key As Variant, ByRef Value As Variant)
    InsertItem Key, Value, False, False
End Property

''
' Gets an <a href="ICollection.html">ICollection</a> containing the keys in the <a href="Hashtable.html">Hashtable</a>.
'
' @return An <a href="ICollection.html">ICollection</a> containing the keys in the <a href="Hashtable.html">Hashtable</a>.
' @remarks The order of the keys in the <a href="ICollection.html">ICollection</a> is unspecified, but it is the same order
' as the associated values in the <a href="ICollection.html">ICollection</a> returned by the <a href="Hashtable.Get.Values.html">Values</a> method.
' <p>The returned <a href="ICollection.html">ICollection</a> is not a static copy; instead, the <a href="ICollection.html">ICollection</a> refers back
' to the keys in the original <a href="Hashtable.html">Hashtable</a>. Therefore, changes to the <a href="Hashtable.html">Hashtable</a> continue to be reflected in the <a href="ICollection.html">ICollection</a>.</p>
' @see #Values
Public Property Get Keys() As ICollection
    Dim Ret As New HashtableKeyCollection
    Ret.Init Me
    Set Keys = Ret
End Property

''
' Returns an enumerator for the table.
'
' @return The enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Removes the element with the specified key from the <a href="Hashtable.html">Hashtable</a>.
'
' @param Key The key of the element to remove.
' @see #Add
' @include "..\DocIncludes\Hashtable.Remove.html"
Public Sub Remove(ByRef Key As Variant)
    Dim Index As Long
    Index = GetKeyIndex(Key)
    If Index >= 0 Then
        With mBuckets(Index)
            .HashCode = 0
            .Key = Empty
            .State = DeletedState
            .Value = Empty
        End With
        
        mCount = mCount - 1
        
        ' If there are no more items, we want to ensure all buckets
        ' are marked as empty for faster instertions.
        If mCount = 0 Then
            Dim i As Long
            For i = 0 To UBound(mBuckets)
                mBuckets(i).State = EmptyState
            Next
        End If
        
        mVersion = mVersion + 1
    End If
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Gets an <a href="ICollection.html">ICollection</a> containing the values in the <a href="Hashtable.html">Hashtable</a>.
'
' @return An <a href="ICollection.html">ICollection</a> containing the values in the <a href="Hashtable.html">Hashtable</a>.
' @remarks The order of the values in the <a href="ICollection.html">ICollection</a> is unspecified, but it is the same order
' as the associated keys in the <a href="ICollection.html">ICollection</a> returned by the <a href="Hashtable.Get.Keys.html">Keys</a> method.
' <p>The returned <a href="ICollection.html">ICollection</a> is not a static copy; instead, the <a href="ICollection.html">ICollection</a> refers back
' to the values in the original <a href="Hashtable.html">Hashtable</a>. Therefore, changes to the <a href="Hashtable.html">Hashtable</a> continue to be reflected in the <a href="ICollection.html">ICollection</a>.</p>
' @see #Keys
Public Property Get Values() As ICollection
    Dim Ret As New HashtableValueCollection
    Ret.Init Me
    Set Values = Ret
End Property

Friend Sub CopyKeys(ByRef Arr As Variant, ByVal ArrayIndex As Long)
    CopyToCore Arr, ArrayIndex, EnumKeys
End Sub

Friend Sub CopyValues(ByRef Arr As Variant, ByVal ArrayIndex As Long)
    CopyToCore Arr, ArrayIndex, EnumValues
End Sub

Friend Function GetKeyEnumerator() As IEnumerator
    Dim Ret As New HashtableEnumerator
    Ret.Init Me, mBuckets, EnumKeys
    Set GetKeyEnumerator = Ret
End Function

Friend Function GetValueEnumerator() As IEnumerator
    Dim Ret As New HashtableEnumerator
    Ret.Init Me, mBuckets, EnumValues
    Set GetValueEnumerator = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Init Nothing, 0, 1!, Nothing
End Sub

Private Sub Class_InitProperties()
    Init Nothing, 0, 1!, Nothing
End Sub

Friend Sub Init(ByVal d As IDictionary, ByVal Capacity As Long, ByVal LoadFactor As Single, ByVal EqualityComparer As IEqualityComparer)
    mLoadFactor = LoadFactor * 0.72!
    
    If EqualityComparer Is Nothing Then
        Set mKeyComparer = Statics.EqualityComparer.Default
    Else
        Set mKeyComparer = EqualityComparer
    End If
    
    If d Is Nothing Then
        If Capacity > mCapacity Or mCapacity = 0 Then
            SetCapacity Capacity
            ReDim mBuckets(0 To mCapacity - 1)
        End If
    Else
        SetCapacity d.Count
        ReDim mBuckets(0 To mCapacity - 1)
        
        With d.GetEnumerator
            Do While .MoveNext
                Add .Key, .Value
            Loop
        End With
    End If
End Sub

Friend Sub InitClone(ByRef Buckets() As HashBucket, ByVal Count As Long, ByVal LoadFactor As Single, ByVal LoadSize As Long, ByVal KeyComparer As IEqualityComparer)
    mCount = Count
    mBuckets = Buckets
    mCapacity = UBound(Buckets) + 1
    mLoadFactor = LoadFactor
    mLoadSize = LoadSize
    Set mKeyComparer = KeyComparer
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCapacity(ByVal Capacity As Long)
    mCapacity = GetHashtableCapacity(Capacity)
    mLoadSize = mCapacity * mLoadFactor
End Sub

Private Sub Expand()
    SetCapacity GetHashtableCapacity(mCapacity * 2)
    ReinsertItems
End Sub

Private Sub ReinsertItems()
    Dim OldBuckets() As HashBucket
    SAPtr(OldBuckets) = SAPtr(mBuckets)
    SAPtr(mBuckets) = vbNullPtr
    
    ReDim mBuckets(0 To mCapacity - 1)
    mCount = 0
    
    Dim i As Long
    For i = 0 To UBound(OldBuckets)
        With OldBuckets(i)
            If .State = OccupiedState Then
                InsertItem .Key, .Value, True, True
            End If
        End With
    Next i
End Sub

Private Function GetKeyIndex(ByRef Key As Variant) As Long
    Dim Step        As Long
    Dim Index       As Long
    Dim HashCode    As Long
    Dim FirstIndex  As Long
    
    HashCode = InitHashCode(Key, Step, Index)
    FirstIndex = Index
    
    Do
        With mBuckets(Index)
            Select Case .State
                Case EmptyState
                    GetKeyIndex = -1
                    Exit Function
                Case OccupiedState
                    If .HashCode = HashCode Then
                        If mKeyComparer.Equals(Key, .Key) Then
                            GetKeyIndex = Index
                            Exit Function
                        End If
                    End If
            End Select
        End With
        
        Index = (Index + Step) Mod mCapacity
    Loop While Index <> FirstIndex
    
    GetKeyIndex = -1
End Function

Private Function InitHashCode(ByRef Key As Variant, ByRef Step As Long, ByRef Index As Long) As Long
    InitHashCode = mKeyComparer.GetHashCode(Key) And &H7FFFFFFF
    Step = 1 + (((InitHashCode \ 32) + 1) Mod (mCapacity - 1))
    Index = InitHashCode Mod mCapacity
End Function

Private Sub InsertItem(ByRef Key As Variant, ByRef Value As Variant, ByVal Adding As Boolean, ByVal Steal As Boolean)
    Dim Step                As Long
    Dim Index               As Long
    Dim FirstEmpty          As Long
    Dim HashCode            As Long
    Dim FirstIndex          As Long
    
    If mCount > mLoadSize Then
        Expand
    End If
    
    HashCode = InitHashCode(Key, Step, Index)
    FirstIndex = Index
    FirstEmpty = -1
    
    Do
        With mBuckets(Index)
            Select Case .State
                Case EmptyState
                    ' we bookmark this bucket as our first available bucket.
                    If FirstEmpty = -1 Then
                        FirstEmpty = Index
                    End If
                    
                    ' we can exit here because we have found an empty bucket.
                    ' Meaning there are no more buckets on this chain so no
                    ' duplicate key could exist.
                    Exit Do
                    
                Case OccupiedState
                    If .HashCode = HashCode Then
                        If mKeyComparer.Equals(Key, .Key) Then
                            If Adding Then _
                                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_AddingDuplicate_Key, .Key, Key), "key")
                                    
                            VariantCopyInd .Key, Key
                            VariantCopyInd .Value, Value
                            mVersion = mVersion + 1
                            Exit Sub
                        End If
                    End If
                
                Case DeletedState
                    ' we bookmark this bucket as our first available bucket
                    ' to place the value and key in. We only do this once
                    ' to the first occurence of a deleted bucket.
                    If FirstEmpty = -1 Then
                        FirstEmpty = Index
                    End If
                    
                    ' we don't exit here because we need to keep
                    ' checking to see if the key exists in the
                    ' chain of buckets further down.
            End Select
        End With
        
        Index = (Index + Step) Mod mCapacity
        
        ' If we end up on the starting index then we have
        ' circled back around, so exit or we will be in an infinite loop.
    Loop While Index <> FirstIndex
    
    Debug.Assert FirstEmpty > -1    ' this should never happen. Check loadfactor.
    If FirstEmpty = -1 Then _
        Throw Cor.NewInvalidOperationException("Hashtable insert failed.")
        
    With mBuckets(FirstEmpty)
        If Steal Then
            Helper.MoveVariant .Key, Key
            Helper.MoveVariant .Value, Value
        Else
            VariantCopyInd .Key, Key
            VariantCopyInd .Value, Value
        End If
        .State = OccupiedState
        .HashCode = HashCode
    End With
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub CopyToCore(ByRef Arr As Variant, ByVal ArrayIndex As Long, ByVal EnumType As EnumerationType)
    If CorArray.IsNull(Arr) Then _
        Error.ArgumentNull "Arr", ArgumentNull_Array
    If CorArray.Rank(Arr) <> 1 Then _
        Error.Argument Arg_RankMultiDimNotSupported, "Arr"
    If ArrayIndex < LBound(Arr) Then _
        Error.ArgumentOutOfRange "ArrayIndex", ArgumentOutOfRange_LBound
    If ArrayIndex + mCount > Len1D(Arr) Then _
        Error.Argument Arg_ArrayPlusOffTooSmall
    
    If mCount = 0 Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(mBuckets)
        With mBuckets(i)
            If .State = OccupiedState Then
                Select Case EnumType
                    Case EnumEntries
                        Set Arr(ArrayIndex + j) = Cor.NewDictionaryEntry(.Key, .Value)
                    Case EnumKeys
                        If IsObject(.Key) Then
                            Set Arr(ArrayIndex + j) = .Key
                        Else
                            Arr(ArrayIndex + j) = .Key
                        End If
                    Case EnumValues
                        If IsObject(.Value) Then
                            Set Arr(ArrayIndex + j) = .Value
                        Else
                            Arr(ArrayIndex + j) = .Value
                        End If
                End Select
                                
                j = j + 1
                
                If j = mCount Then
                    Exit Sub
                End If
            End If
        End With
    Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Dim Comparer As IEqualityComparer
        If .ReadProperty(PropUseDefaultComparer, False) Then
            Set Comparer = EqualityComparer.Default
        Else
            Set Comparer = .ReadProperty(PropComparer, Nothing)
        End If
        
        Dim Count As Long
        Count = .ReadProperty(PropCount, 0)
        Init Nothing, Count, DefaultLoadFactor, Comparer
        
        Dim i As Long
        For i = 1 To Count
            ' When persisting data within a Class_WriteProperties event, it seems that
            ' object values of Nothing and variant values of Empty do not get persisted
            ' correctly, or at all, except the key name. Even when reading with a default
            ' value of Empty, it returns an empty string "". So now we maintain a flag
            ' with each value to determine the correct value.
            Select Case .ReadProperty(MakeValueSubTypeKey(i), SubtypeNormal)
                Case SubtypeNormal
                    Add .ReadProperty(MakeKeyKey(i), Nothing), .ReadProperty(MakeValueKey(i))
                Case SubtypeEmpty
                    Add .ReadProperty(MakeKeyKey(i), Nothing), Empty
                Case SubtypeNothing
                    Add .ReadProperty(MakeKeyKey(i), Nothing), Nothing
            End Select
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        ' Simply save the number of elements are in this Hashtable.
        .WriteProperty PropCount, mCount
        
        ' We don't want to create a duplicate of the default comparer object,
        ' so we will just set a flag that says to use the default comparer
        ' object when the Hashtable is deserialized.
        If mKeyComparer Is EqualityComparer.Default Then
            .WriteProperty PropUseDefaultComparer, True
        Else
            ' Either we aren't using a comparer at all or we are
            ' using a user-supplied comparer. If so, we hope it
            ' is persistable, or an exception will be thrown.
            .WriteProperty PropComparer, mKeyComparer
        End If
        
        ' Iterate through the buckets in this hashtable. If the bucket
        ' is marked 'occupied' then we will save the key and value.
        ' We hope the key and value are persistable. If not, then
        ' an exception will be thrown.
        Dim i       As Long
        Dim Count   As Long
        For i = 0 To mCapacity - 1
            With mBuckets(i)
                If .State = OccupiedState Then
                    Count = Count + 1
                    PropBag.WriteProperty MakeKeyKey(Count), .Key
                    
                    ' When persisting data within a Class_WriteProperties event, it seems that
                    ' object values of Nothing and variant values of Empty do not get persisted
                    ' correctly, or at all, except the key name. Even when reading with a default
                    ' value of Empty, it returns an empty string "". So now we maintain a flag
                    ' with each value to determine the correct value.
                    Select Case VarType(.Value)
                        Case vbEmpty
                            PropBag.WriteProperty MakeValueSubTypeKey(Count), SubtypeEmpty
                        
                        Case vbObject
                            If .Value Is Nothing Then
                                PropBag.WriteProperty MakeValueSubTypeKey(Count), SubtypeNothing
                            Else
                                PropBag.WriteProperty MakeValueKey(Count), .Value
                            End If
                        
                        Case Else
                            PropBag.WriteProperty MakeValueKey(Count), .Value
                    
                    End Select
                End If
            End With
        Next i
    End With
End Sub

Private Function MakeKeyKey(ByVal Count As Long) As String
    MakeKeyKey = PropKey & Count
End Function

Private Function MakeValueKey(ByVal Count As Long) As String
    MakeValueKey = PropValue & Count
End Function

Private Function MakeValueSubTypeKey(ByVal Count As Long) As String
    MakeValueSubTypeKey = PropValueSubtype & Count
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
