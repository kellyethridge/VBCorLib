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
' The Hashtable represents a repository of key/value pairs.
'
' @remarks This class is persistable.
'
' @see Constructors
' @see IDictionary
' @see ICollection
' @see IEnumerable
' @see ICloneable
' @see CaseInsensitiveHashCodePrvdr
' @see CaseInsensitiveComparer
'
Option Explicit
Implements IObject
Implements IDictionary
Implements ICollection
Implements IEnumerable
Implements ICloneable
Implements IVersionableCollection

Private Const PropCount                 As String = "Count"
Private Const PropKey                   As String = "Key"
Private Const PropValueSubtype          As String = "ValueSubType"
Private Const PropValue                 As String = "Value"
Private Const PropProvider              As String = "Provider"
Private Const PropComparer              As String = "Comparer"
Private Const PropUseDefaultComparer    As String = "UseDefaultComparer"
Private Const SubtypeNormal             As Long = 0
Private Const SubtypeEmpty              As Long = 1
Private Const SubtypeNothing            As Long = 2
Private Const DefaultCapacity           As Long = 32
Private Const DefaultLoadFactor         As Single = 1!

Private mLoadFactor     As Single
Private mLoadSize       As Long
Private mCount          As Long
Private mBuckets()      As HashBucket
Private mCapacity       As Long
Private mVersion        As Long
Private mKeyComparer    As IEqualityComparer


''
' Adds a new key/value pair to the table.
'
' @param Key The key used to identify the value.
' @param Value The value to be added to the table.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object does not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.
'
Public Sub Add(ByRef Key As Variant, ByRef Value As Variant)
    InsertItem Key, Value, True, False
End Sub

''
' Clears all of the elements from the table.
'
' @remarks When clearing the table, the count is set to 0, but the
' capacity remains unchanged.
'
Public Sub Clear()
    ReDim mBuckets(0 To mCapacity - 1)
    mCount = 0
    mVersion = mVersion + 1
End Sub

''
' Causes the Hashtable to clean up deleted buckets to improve
' searching capabilities.
'
' @remarks When many items have been deleted from the table, the
' buckets that contained the items are marked as Deleted. When many
' buckets are marked as Deleted, the search routine cannot stop
' searching the buckets for a duplicate key even if an empty Deleted
' bucket has been found. So to prevent the unnessecary searching past
' Deleted buckets, the table should be cleaned to remove the Deleted
' buckets and only have Empty and Occupied buckets remaining.
' <p>This is recommended when a large number of items are deleted.</p>
'
Public Sub Clean()
    ReinsertItems
    mVersion = mVersion + 1
End Sub

''
' Returns a clone of the current instance.
'
' @return The clone of the current instance.
'
Public Function Clone() As Hashtable
    Set Clone = New Hashtable
    Clone.InitClone mBuckets, mCount, mLoadFactor, mLoadSize, mKeyComparer
End Function

''
' Checks if the current instance contains a specific key.
'
' @param Key The key to check for.
' @return Indicates if the key was found.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.
'
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = ContainsKey(Key)
End Function

''
' Checks if the current instance contains a specific key.
'
' @param Key The key to check for.
' @return Indicates if the key was found.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.
'
Public Function ContainsKey(ByRef Key As Variant) As Boolean
    ContainsKey = (GetKeyIndex(Key) >= 0)
End Function

''
' Checks if the current instance contains a specific value.
'
' @param Value The value to search for.
' @return Indicates if the value was found.
' @remarks The table is searched linearly checking each value. Once
' a matching value is found, True is returned. If there is more than
' one of the same values, only the first is detected.
'
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
' Copies the contents of the table to an array.
'
' @param Arr The array to copy the contents to.
' @param ArrayIndex The starting index in dstArray to begin copying to.
' @remarks The elements that are copied to the array are DictionaryEntry
' elements. Each DictionaryEntry contains a key and value from the table.
'
Public Sub CopyTo(ByRef Arr As Variant, ByVal ArrayIndex As Long)
    CopyToCore Arr, ArrayIndex, EnumEntries
End Sub

''
' Returns the number of elements in the table.
'
' @return The number of elements in the table.
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
' Returns an enumerator for the table.
'
' @return An IEnumerator object that enumerates over the table.
' @remarks The values enumerated are DictionaryEntry object each
' containing a key/value pair.
'
Public Function GetEnumerator() As Object
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
' Returns if this instance is fixed size.
'
' @return Default is False.
' @remarks This is to satisfy the IList interface.
'
Public Property Get IsFixedSize() As Boolean
    IsFixedSize = False
End Property

''
' Returns if this instance is readonly
'
' @return Default is False.
' @remarks This is to satisfy the IList interface.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns a value associated with the specified key.
'
' @param Key The key used to look up a value.
' @return The value associated with the Key.
' @remarks If the key does not exist, then Empty is returned.
' <p>The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.</p>
'
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Index As Long
    Index = GetKeyIndex(Key)
    If Index >= 0 Then
        VariantCopy Item, mBuckets(Index).Value
    End If
End Property

''
' Sets the value associated to the specified key.
'
' @param Key The key used to look up a value.
' @param Value The value to associate with the key.
' @remarks If the key is not found, then the key/value are added to
' the table. If the key IS found, then the associated value with the
' key is replaced with the new value.
' <p>The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.</p>
'
Public Property Let Item(ByRef Key As Variant, ByRef Value As Variant)
    InsertItem Key, Value, False, False
End Property

''
' Sets the value associated to the specified key.
'
' @param Key The key used to look up a value.
' @param Value The value to associate with the key.
' @remarks If the key is not found, then the key/value are added to
' the table. If the key IS found, then the associated value with the
' key is replaced with the new value.
' <p>The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object does not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.</p>
'
Public Property Set Item(ByRef Key As Variant, ByRef Value As Variant)
    InsertItem Key, Value, False, False
End Property

''
' Retuns an ICollection object used to access the keys of the table.
'
' @return An ICollection object used to access the keys of the table.
'
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
' Removes a key/value pair associated with a specific key.
'
' @param Key The key use to lookup the key/value pair to be removed.
'
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
' Retuns an ICollection object used to access the values of the table.
'
' @return An ICollection object used to access the values of the table.
'
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
Private Property Get IVersionableCollection_CurrentVersion() As Long
    IVersionableCollection_CurrentVersion = mVersion
End Property
