VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedListEnumerator"
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
' Module: SortedListEnumerator
'
Option Explicit
Implements IObject
Implements IEnumerator
Implements ICloneable

Public Enum SortedListEnumType
    slEntries
    slKeys
    slValues
End Enum


Private mList       As SortedList
Private mIndex      As Long
Private mEndIndex   As Long
Private mKeys()     As Variant
Private mValues()   As Variant
Private mVersion    As Long
Private mEnumType   As SortedListEnumType


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As SortedList, ByVal KeysSAPtr As Long, ByVal ValuesSAPtr As Long, ByVal EnumType As SortedListEnumType)
    Set mList = List
    SAPtr(mKeys) = KeysSAPtr
    SAPtr(mValues) = ValuesSAPtr
    mVersion = List.Version
    mEndIndex = List.Count - 1
    mIndex = -1
    mEnumType = EnumType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    SAPtr(mKeys) = 0
    SAPtr(mValues) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim Ret As New SortedListEnumerator
    Call Ret.Init(mList, SAPtr(mKeys), SAPtr(mValues), mEnumType)
    Set ICloneable_Clone = Ret
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex = -1 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    Select Case mEnumType
        Case slEntries
            Set IEnumerator_Current = Cor.NewDictionaryEntry(mKeys(mIndex), mValues(mIndex))
        Case slKeys
            Call VariantCopy(IEnumerator_Current, mKeys(mIndex))
        Case slValues
            Call VariantCopy(IEnumerator_Current, mValues(mIndex))
    End Select
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
    
    If mIndex = mEndIndex Then
        mIndex = mEndIndex + 1
    Else
        mIndex = mIndex + 1
        IEnumerator_MoveNext = True
    End If
End Function

Private Sub IEnumerator_Reset()
    mIndex = -1
End Sub

