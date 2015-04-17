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
'    Module: HashtableEnumerator
'
''
' Class derived from IObject.
'

Option Explicit
Implements IObject
Implements IDictionaryEnumerator
Implements IEnumerator

Private mList As Hashtable
Private mBuckets() As Bucket
Private mEnumType As IDictionaryEnumType
Private mIndex As Long
Private mVersion As Long
Private mCount As Long
Private mEntries As Long
Private mCapacity As Long


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function

Friend Sub Init(ByVal List As Hashtable, ByVal BucketsSAPtr As Long, ByVal HTEnumType As IDictionaryEnumType)
    Set mList = List
    SAPtr(mBuckets) = BucketsSAPtr
    mEnumType = HTEnumType
    mIndex = -1
    mVersion = List.Version
    mCount = List.Count
    mCapacity = UBound(mBuckets) + 1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyValidEnum()
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    SAPtr(mBuckets) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IDictionaryEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IDictionaryEnumerator_Current() As Variant
    Call VerifyValidEnum
    
    Select Case mEnumType
        Case detEntries
            Set IDictionaryEnumerator_Current = Cor.NewDictionaryEntry(mBuckets(mIndex).Key, mBuckets(mIndex).Value)
        Case detKeys
            VariantCopy IDictionaryEnumerator_Current, mBuckets(mIndex).Key
        Case detValues
            VariantCopy IDictionaryEnumerator_Current, mBuckets(mIndex).Value
    End Select
End Property

Private Property Get IDictionaryEnumerator_Entry() As DictionaryEntry
    Call VerifyValidEnum
    
    Set IDictionaryEnumerator_Entry = Cor.NewDictionaryEntry(mBuckets(mIndex).Key, mBuckets(mIndex).Value)
End Property

Private Property Get IDictionaryEnumerator_Key() As Variant
    Call VerifyValidEnum
    Call VariantCopy(IDictionaryEnumerator_Key, mBuckets(mIndex).Key)
End Property

Private Function IDictionaryEnumerator_MoveNext() As Boolean
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mEntries > mCount Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
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
    Call VerifyValidEnum
    Call VariantCopy(IDictionaryEnumerator_Value, mBuckets(mIndex).Value)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    Call Helper.MoveVariant(IEnumerator_Current, IDictionaryEnumerator_Current)
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = IDictionaryEnumerator_MoveNext
End Function

Private Sub IEnumerator_Reset()
    IDictionaryEnumerator_Reset
End Sub
