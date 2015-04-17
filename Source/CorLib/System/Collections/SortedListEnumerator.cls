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
'    Module: SortedListEnumerator
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

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
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

