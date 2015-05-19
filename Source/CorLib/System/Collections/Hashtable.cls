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
Private Const LOAD_FACTOR   As Double = 0.9

Private mCount          As Long
Private mBuckets()      As Bucket
Private mCapacity       As Long
Private mVersion        As Long
Private mLoadThreshold  As Long
Private mProvider       As IHashcodeProvider
Private mComparer       As IComparer

' used in hashing a string key.
Private mStringHashChars As WordBuffer



''
' Adds a new key/value pair to the table.
'
' @param Key The key used to identify the value.
' @param Value The value to be added to the table.
' @remarks The key can be any datatype other than vbUserDefinedType.
' If an object is being used as a key, then it should implement the
' IObject interface to allow for custom hashcode creation. If the
' object doesn't not implement the interface, then the objects memory
' location is used for a hashcode. If two different objects should
' represent the same hashcode, then they need to implement the IObject
' interface and override the GetHashCode function.
'
Public Sub Add(ByRef Key As Variant, ByRef Value As Variant)
    Call InsertItem(Key, Value, True, False)
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
    Clone.CloneHelper mBuckets, mCount, mLoadThreshold, mProvider, mComparer
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
        If mBuckets(i).State = bsOccupied Then
            If EqualVariants(mBuckets(i).Value, Value) Then
                ContainsValue = True
                Exit Function
            End If
        End If
    Next i
End Function

''
' Copies the contents of the table to an array.
'
' @param dstArray The array to copy the contents to.
' @param index The starting index in dstArray to begin copying to.
' @remarks The elements that are copied to the array are DictionaryEntry
' elements. Each DictionaryEntry contains a key and value from the table.
'
Public Sub CopyTo(ByRef DstArray As Variant, ByVal Index As Long)
    InternalCopyTo DstArray, Index, detEntries
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
    Ret.Init Me, SAPtr(mBuckets), detEntries
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
    Dim i As Long
    i = GetKeyIndex(Key)
    If i >= 0 Then
        VariantCopy Item, mBuckets(i).Value
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
    Dim Ret As New HTKeyCollection
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
    Dim i As Long
    i = GetKeyIndex(Key)
    If i >= 0 Then
        With mBuckets(i)
            .hashcode = 0
            .Key = Empty
            .State = bsDeleted
            .Value = Empty
        End With
        
        mCount = mCount - 1
        
        ' If there are no more items, we want to ensure all buckets
        ' are marked as empty for faster instertions.
        If mCount = 0 Then
            For i = 0 To UBound(mBuckets)
                mBuckets(i).State = bsEmpty
            Next i
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
    Dim Ret As New HTValueCollection
    Ret.Init Me
    Set Values = Ret
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub CloneHelper(ByRef buckets() As Bucket, ByVal Count As Long, LoadThreshold As Long, ByVal Provider As IHashcodeProvider, ByVal Comparer As IComparer)
    mCount = Count
    mBuckets = buckets
    mCapacity = UBound(buckets) + 1
    mLoadThreshold = LoadThreshold
    Set mProvider = Provider
    Set mComparer = Comparer
End Sub

Friend Sub CopyKeys(ByRef DstArray As Variant, ByVal Index As Long)
    InternalCopyTo DstArray, Index, detKeys
End Sub

Friend Sub CopyValues(ByRef DstArray As Variant, ByVal Index As Long)
    InternalCopyTo DstArray, Index, detValues
End Sub

Friend Function GetKeyEnumerator() As IEnumerator
    Dim Ret As New HashtableEnumerator
    Ret.Init Me, SAPtr(mBuckets), detKeys
    Set GetKeyEnumerator = Ret
End Function

Friend Function GetValueEnumerator() As IEnumerator
    Dim Ret As New HashtableEnumerator
    Ret.Init Me, SAPtr(mBuckets), detValues
    Set GetValueEnumerator = Ret
End Function

Friend Sub InitCapacity(ByVal Capacity As Long, ByVal hcp As IHashcodeProvider, ByVal Comparer As IComparer)
    mCapacity = GetPrime(Capacity)
    mLoadThreshold = mCapacity * LOAD_FACTOR
    ReDim mBuckets(0 To mCapacity - 1)
    Set mProvider = hcp
    Set mComparer = Comparer
End Sub

Friend Sub InitDictionary(ByVal dic As IDictionary, ByVal hcp As IHashcodeProvider, ByVal Comparer As IComparer)
    If dic Is Nothing Then
        InitCapacity DEF_CAPACITY, hcp, Comparer
    Else
        InitCapacity dic.Count, hcp, Comparer
        Dim Entry As DictionaryEntry
        For Each Entry In dic
            InsertItem Entry.Key, Entry.Value, True, True
        Next Entry
    End If
End Sub

Friend Property Get Version() As Long
    Version = mVersion
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExpandTable()
    mCapacity = GetPrime(mCapacity * 2)
    mLoadThreshold = mCapacity * LOAD_FACTOR
    ReinsertItems
End Sub

Private Sub ReinsertItems()
    Dim OldBuckets() As Bucket
    ReDim OldBuckets(0 To mCapacity - 1)
    Helper.Swap4 ByVal ArrPtr(OldBuckets), ByVal ArrPtr(mBuckets)
    
    mCount = 0
    
    Dim i As Long
    For i = 0 To UBound(OldBuckets)
        If OldBuckets(i).State = bsOccupied Then
            InsertItem OldBuckets(i).Key, OldBuckets(i).Value, True, True
        End If
    Next i
End Sub

Private Function GetKeyIndex(ByRef Key As Variant) As Long
    Dim Step                As Long
    Dim Index               As Long
    Dim hashcode            As Long
    Dim FirstIndex          As Long
    Dim ComparerIsNothing   As Boolean
    
    hashcode = InitHashCode(Key, Step, Index)
    FirstIndex = Index
    ComparerIsNothing = mComparer Is Nothing
    
    Do
        Select Case mBuckets(Index).State
            Case bsEmpty
                GetKeyIndex = -1
                Exit Function
            Case bsOccupied
                If mBuckets(Index).hashcode = hashcode Then
                    If ComparerIsNothing Then
                        If EqualVariants(Key, mBuckets(Index).Key) Then
                            GetKeyIndex = Index
                            Exit Function
                        End If
                    ElseIf mComparer.Compare(Key, mBuckets(Index).Key) = 0 Then
                        GetKeyIndex = Index
                        Exit Function
                    End If
                End If
        End Select
        Index = (Index + Step) Mod mCapacity
    Loop While Index <> FirstIndex
    
    GetKeyIndex = -1
End Function

Private Function InitHashCode(ByRef Key As Variant, ByRef Step As Long, ByRef Index As Long) As Long
    If mProvider Is Nothing Then
        InitHashCode = Object.GetHashCode(Key) And &H7FFFFFFF
    Else
        InitHashCode = mProvider.GetHashCode(Key) And &H7FFFFFFF
    End If
    Step = 1 + (((InitHashCode \ 32) + 1) Mod (mCapacity - 1))
    Index = InitHashCode Mod mCapacity
End Function

Private Sub InsertItem(ByRef Key As Variant, ByRef Value As Variant, ByVal Adding As Boolean, ByVal Steal As Boolean)
    Dim Step                As Long
    Dim Index               As Long
    Dim FirstEmpty          As Long
    Dim hashcode            As Long
    Dim UseDefaultCompare   As Boolean
    Dim FirstIndex          As Long
    
    If mCount > mLoadThreshold Then Call ExpandTable
    
    hashcode = InitHashCode(Key, Step, Index)
    FirstIndex = Index
    FirstEmpty = -1
    UseDefaultCompare = mComparer Is Nothing
    
    Do
        Select Case mBuckets(Index).State
            Case bsEmpty
                ' we bookmark this bucket as our first available bucket.
                If FirstEmpty = -1 Then FirstEmpty = Index
                
                ' we can exit here because we have found an empty bucket.
                ' Meaning there are no more buckets on this chain so no
                ' duplicate key could exist.
                Exit Do
                
            Case bsOccupied
                If mBuckets(Index).hashcode = hashcode Then
                    If UseDefaultCompare Then
                        If EqualVariants(Key, mBuckets(Index).Key) Then
                            If Adding Then _
                                Throw Cor.NewArgumentException("An element with the same key already exists in the collection.", "key")
                            
                            ' v2.0 - we copy the key so the same updated key stays with the updated value.
                            VariantCopyInd mBuckets(Index).Key, Key
                            VariantCopyInd mBuckets(Index).Value, Value
                            mVersion = mVersion + 1
                            Exit Sub
                        End If
                    ElseIf mComparer.Compare(Key, mBuckets(Index).Key) = 0 Then
                        If Adding Then _
                            Throw Cor.NewArgumentException("An element with the same key already exists in the collection.", "key")
                        
                        ' v2.0 - we copy the key so the same updated key stays with the updated value.
                        VariantCopyInd mBuckets(Index).Key, Key
                        VariantCopyInd mBuckets(Index).Value, Value
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
        .hashcode = hashcode
    End With
    mCount = mCount + 1
    mVersion = mVersion + 1
End Sub

Private Sub InternalCopyTo(ByRef DstArray As Variant, ByVal Index As Long, ByVal CopyType As IDictionaryEnumType)
    Dim Result As Long
    Result = VerifyArrayRange(CorArray.ArrayPointer(DstArray), Index, mCount)
    If Result <> NO_ERROR Then
        ThrowArrayRangeException Result, "DstArray", Index, "Index", mCount, "Index"
    End If
    
    Dim i As Long
    Dim j As Long
    For i = 0 To mCapacity - 1
        If mBuckets(i).State = bsOccupied Then
            Select Case CopyType
                Case detEntries
                    Set DstArray(Index + j) = Cor.NewDictionaryEntry(mBuckets(i).Key, mBuckets(i).Value)
                Case detKeys
                    If IsObject(mBuckets(i).Key) Then
                        Set DstArray(Index + j) = mBuckets(i).Key
                    Else
                        DstArray(Index + j) = mBuckets(i).Key
                    End If
                Case detValues
                    If IsObject(mBuckets(i).Value) Then
                        Set DstArray(Index + j) = mBuckets(i).Value
                    Else
                        DstArray(Index + j) = mBuckets(i).Value
                    End If
            End Select
            j = j + 1
            If j = mCount Then Exit Sub
        End If
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitWordBuffer mStringHashChars, 0, 0
End Sub

Private Sub Class_InitProperties()
    InitCapacity DEF_CAPACITY, Nothing, Nothing
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Dim Comparer As Object
        If .ReadProperty(PROP_USEDEFAULTCOMPARER, False) Then
            Set Comparer = Comparer.Default
        Else
            Set Comparer = .ReadProperty(PROP_COMPARER, Nothing)
        End If
        
        Dim Count As Long
        Count = .ReadProperty(PROP_COUNT)
        InitCapacity Count, .ReadProperty(PROP_PROVIDER, Nothing), Comparer
        
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
        
        ' The provider will either be Nothing or a user-supplied
        ' provider. If so, we hope it is persitable or an exception
        ' will be thrown.
        .WriteProperty PROP_PROVIDER, mProvider
        
        ' We don't want to create a duplicate of the default comparer object,
        ' so we will just set a flag that says to use the default comparer
        ' object when the Hashtable is deserialized.
        If mComparer Is Comparer.Default Then
            .WriteProperty PROP_USEDEFAULTCOMPARER, True
        Else
            ' Either we aren't using a comparer at all or we are
            ' using a user-supplied comparer. If so, we hope it
            ' is persistable, or an exception will be thrown.
            .WriteProperty PROP_COMPARER, mComparer
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
'   IDictionary Interface
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function
