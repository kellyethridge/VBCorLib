VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashtableEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: HashtableEnumerator
'
Option Explicit
Implements IObject
Implements IDictionaryEnumerator
Implements IEnumerator
Implements ICloneable

Private mTable      As IVersionableCollection
Private mBuckets()  As Bucket
Private mEnumType   As EnumerationType
Private mIndex      As Long
Private mVersion    As Long
Private mCount      As Long
Private mEntries    As Long
Private mCapacity   As Long


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyValidEnum()
    If mTable.CurrentVersion <> mVersion Then
        Error.InvalidOperation InvalidOperation_VersionError
    End If
    If mIndex < 0 Then
        Error.InvalidOperation InvalidOperation_EnumNotStarted
    End If
    If mEntries > mCount Then
        Error.InvalidOperation InvalidOperation_EnumFinished
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Table As Hashtable, ByRef Buckets() As Bucket, ByVal EnumType As EnumerationType)
    Set mTable = Table
    SAPtr(mBuckets) = SAPtr(Buckets)
    mEnumType = EnumType
    mIndex = -1
    mVersion = mTable.CurrentVersion
    mCount = Table.Count
    mCapacity = UBound(mBuckets) + 1
End Sub

Friend Sub InitClone(ByVal Table As Hashtable, ByRef Buckets() As Bucket, ByVal EnumType As EnumerationType, ByVal Index As Long, ByVal Entries As Long, ByVal Capacity As Long)
    Set mTable = Table
    mVersion = mTable.CurrentVersion
    SAPtr(mBuckets) = SAPtr(Buckets)
    mEnumType = EnumType
    mIndex = Index
    mCount = Table.Count
    mEntries = Entries
    mCapacity = Capacity
End Sub

Private Sub Class_Terminate()
    SAPtr(mBuckets) = vbNullPtr
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim Ret As New HashtableEnumerator
    Ret.InitClone mTable, mBuckets, mEnumType, mIndex, mEntries, mCapacity
    Set ICloneable_Clone = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    VerifyValidEnum
    
    Select Case mEnumType
        Case EnumEntries
            Set IDictionaryEnumerator_Current = Cor.NewDictionaryEntry(mBuckets(mIndex).Key, mBuckets(mIndex).Value)
        Case EnumKeys
            VariantCopy IDictionaryEnumerator_Current, mBuckets(mIndex).Key
        Case EnumValues
            VariantCopy IDictionaryEnumerator_Current, mBuckets(mIndex).Value
    End Select
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    VerifyValidEnum
    Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(mBuckets(mIndex).Key, mBuckets(mIndex).Value)
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    VerifyValidEnum
    VariantCopy IDictionaryEnumerator_Key, mBuckets(mIndex).Key
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    If GetCollectionVersion(mTable) <> mVersion Then _
        Error.InvalidOperation InvalidOperation_VersionError
    If mEntries > mCount Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
    
    Do
        mIndex = mIndex + 1
        If mIndex = mCapacity Then
            mEntries = mCount + 1
            Exit Function
        End If
    Loop While mBuckets(mIndex).State <> bsOccupied
    mEntries = mEntries + 1
    IDictionaryEnumerator_MoveNext = True
End Function

Private Sub IDictionaryEnumerator_Reset()
    mIndex = -1
    mEntries = 0
End Sub

Private Property Get IDictionaryEnumerator_Value() As Variant
    VerifyValidEnum
    VariantCopy IDictionaryEnumerator_Value, mBuckets(mIndex).Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Helper.MoveVariant IEnumerator_Current, IDictionaryEnumerator_Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = IDictionaryEnumerator_MoveNext
End Function

Private Sub IEnumerator_Reset()
    IDictionaryEnumerator_Reset
End Sub
