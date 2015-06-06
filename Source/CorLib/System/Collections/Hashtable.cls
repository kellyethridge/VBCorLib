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
'    Module: Hashtable
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

Private Const PROP_COUNT                As String = "Count"
Private Const PROP_KEY                  As String = "Key"
Private Const PROP_VALUESUBTYPE         As String = "ValueSubType"
Private Const PROP_VALUE                As String = "Value"
Private Const PROP_PROVIDER             As String = "Provider"
Private Const PROP_COMPARER             As String = "Comparer"
Private Const PROP_USEDEFAULTCOMPARER   As String = "UseDefaultComparer"

Private Const SUBTYPE_NORMAL    As Long = 0
Private Const SUBTYPE_EMPTY     As Long = 1
Private Const SUBTYPE_NOTHING   As Long = 2

Private Const DEF_CAPACITY  As Long = 32
Private Const DefaultLoadFactor As Single = 1!

Private mLoadFactor     As Single
Private mLoadSize       As Long
Private mCount          As Long
Private mBuckets()      As Bucket
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
            If .State = bsOccupied Then
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
' @param Index The starting index in dstArray to begin copying to.
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
            .State = bsDeleted
            .Value = Empty
        End With
        
        mCount = mCount - 1
        
        ' If there are no more items, we want to ensure all buckets
        ' are marked as empty for faster instertions.
        If mCount = 0 Then
            Dim i As Long
            For i = 0 To UBound(mBuckets)
                mBuckets(i).State = bsEmpty
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
        SetCapacity Capacity
        ReDim mBuckets(0 To mCapacity - 1)
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

Friend Sub InitClone(ByRef Buckets() As Bucket, ByVal Count As Long, ByVal LoadFactor As Single, ByVal LoadSize As Long, ByVal KeyComparer As IEqualityComparer)
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
    mCapacity = GetPrime(Capacity)
    mLoadSize = mCapacity * mLoadFactor
End Sub

Private Sub Expand()
    SetCapacity GetPrime(mCapacity * 2)
    ReinsertItems
End Sub

Private Sub ReinsertItems()
    Dim OldBuckets() As Bucket
    SAPtr(OldBuckets) = SAPtr(mBuckets)
    SAPtr(mBuckets) = vbNullPtr
    
    ReDim mBuckets(0 To mCapacity - 1)
    mCount = 0
    
    Dim i As Long
    For i = 0 To UBound(OldBuckets)
        With OldBuckets(i)
            If .State = bsOccupied Then
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
                Case bsEmpty
                    GetKeyIndex = -1
                    Exit Function
                Case bsOccupied
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
                Case bsEmpty
                    ' we bookmark this bucket as our first available bucket.
                    If FirstEmpty = -1 Then FirstEmpty = Index
                    
                    ' we can exit here because we have found an empty bucket.
                    ' Meaning there are no more buckets on this chain so no
                    ' duplicate key could exist.
                    Exit Do
                    
                Case bsOccupied
                    If .HashCode = HashCode Then
                        If mKeyComparer.Equals(Key, .Key) Then
                            If Adding Then _
                                Throw Cor.NewArgumentException("An element with the same key already exists in the collection.", "key")
                                    
                            VariantCopyInd .Key, Key
                            VariantCopyInd .Value, Value
                            mVersion = mVersion + 1
                            Exit Sub
                        End If
                    End If
                
                Case bsDeleted
                    ' we bookmark this bucket as our first available bucket
                    ' to place the value and key in. We only do this once
                    ' to the first occurence of a deleted bucket.
                    If FirstEmpty = -1 Then FirstEmpty = Index
                    
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
        .State = bsOccupied
        .HashCode = HashCode
    End With
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub CopyToCore(ByRef Arr As Variant, ByVal ArrayIndex As Long, ByVal EnumType As EnumerationType)
    Validation.CheckValidSingleDimArray Arr
    
    If ArrayIndex < LBound(Arr) Then _
        Throw Error.ArgumentOutOfRange("ArrayIndex", ArgumentOutOfRange_LBound)
    If ArrayIndex + mCount > CorArray.LengthFirstDim(Arr) Then _
        Error.Argument Arg_ArrayPlusOffTooSmall
    
    If mCount = 0 Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    For i = 0 To UBound(mBuckets)
        With mBuckets(i)
            If .State = bsOccupied Then
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
        If .ReadProperty(PROP_USEDEFAULTCOMPARER, False) Then
            Set Comparer = EqualityComparer.Default
        Else
            Set Comparer = .ReadProperty(PROP_COMPARER, Nothing)
        End If
        
        Dim Count As Long
        Count = .ReadProperty(PROP_COUNT)
        Init Nothing, Count, DefaultLoadFactor, Comparer
        
        Dim i As Long
        For i = 1 To Count
            ' When persisting data within a Class_WriteProperties event, it seems that
            ' object values of Nothing and variant values of Empty do not get persisted
            ' correctly, or at all, except the key name. Even when reading with a default
            ' value of Empty, it returns an empty string "". So now we maintain a flag
            ' with each value to determine the correct value.
            Select Case .ReadProperty("ValueSubType" & i, SUBTYPE_NORMAL)
                Case SUBTYPE_NORMAL
                    Add .ReadProperty(PROP_KEY & i, Nothing), .ReadProperty(PROP_VALUE & i)
                Case SUBTYPE_EMPTY
                    Add .ReadProperty(PROP_KEY & i, Nothing), Empty
                Case SUBTYPE_NOTHING
                    Add .ReadProperty(PROP_KEY & i, Nothing), Nothing
            End Select
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        ' Simply save the number of elements are in this Hashtable.
        .WriteProperty PROP_COUNT, mCount
        
        ' We don't want to create a duplicate of the default comparer object,
        ' so we will just set a flag that says to use the default comparer
        ' object when the Hashtable is deserialized.
        If mKeyComparer Is EqualityComparer.Default Then
            .WriteProperty PROP_USEDEFAULTCOMPARER, True
        Else
            ' Either we aren't using a comparer at all or we are
            ' using a user-supplied comparer. If so, we hope it
            ' is persistable, or an exception will be thrown.
            .WriteProperty PROP_COMPARER, mKeyComparer
        End If
        
        ' Iterate through the buckets in this hashtable. If the bucket
        ' is marked 'occupied' then we will save the key and value.
        ' We hope the key and value are persistable. If not, then
        ' an exception will be thrown.
        Dim i       As Long
        Dim Count   As Long
        For i = 0 To mCapacity - 1
            With mBuckets(i)
                If .State = bsOccupied Then
                    Count = Count + 1
                    PropBag.WriteProperty PROP_KEY & Count, .Key
                    
                    ' When persisting data within a Class_WriteProperties event, it seems that
                    ' object values of Nothing and variant values of Empty do not get persisted
                    ' correctly, or at all, except the key name. Even when reading with a default
                    ' value of Empty, it returns an empty string "". So now we maintain a flag
                    ' with each value to determine the correct value.
                    Select Case VarType(.Value)
                        Case vbEmpty
                            PropBag.WriteProperty PROP_VALUESUBTYPE & Count, SUBTYPE_EMPTY
                        
                        Case vbObject
                            If .Value Is Nothing Then
                                PropBag.WriteProperty PROP_VALUESUBTYPE & Count, SUBTYPE_NOTHING
                            Else
                                PropBag.WriteProperty PROP_VALUE & Count, .Value
                            End If
                        
                        Case Else
                            PropBag.WriteProperty PROP_VALUE & Count, .Value
                    
                    End Select
                End If
            End With
        Next i
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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

Private Property Set IDictionary_Item(Key As Variant, RHS As Variant)
    Set Item(Key) = RHS
End Property

Private Property Let IDictionary_Item(Key As Variant, RHS As Variant)
    Item(Key) = RHS
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
