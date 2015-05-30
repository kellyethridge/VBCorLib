VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArrayEnumerator"
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
'    Module: BitArrayEnumerator
'

''
' This class is used to enumerate the individual bits being
' managed by a BitArray object.
'
Option Explicit
Implements IObject
Implements IEnumerator
Implements ICloneable

Private mList       As BitArray
Private mIndex      As Long
Private mEndIndex   As Long
Private mVersion    As Long



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As BitArray)
    Set mList = List
    mVersion = GetCollectionVersion(List)
    mIndex = -1
    mEndIndex = List.Count - 1
End Sub

Friend Sub InitClone(ByVal List As BitArray, ByVal Index As Long)
    Init List
    mIndex = Index
End Sub


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
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim Ret As New BitArrayEnumerator
    Ret.InitClone mList, mIndex
    Set ICloneable_Clone = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If GetCollectionVersion(mList) <> mVersion Then _
        Throw Error.InvalidOperation(InvalidOperation_VersionError)
    If mIndex < 0 Then _
        Throw Error.InvalidOperation(InvalidOperation_EnumNotStarted)
    If mIndex > mEndIndex Then _
        Throw Error.InvalidOperation(InvalidOperation_EnumFinished)

    IEnumerator_Current = mList(mIndex)
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If GetCollectionVersion(mList) <> mVersion Then _
        Throw Error.InvalidOperation(InvalidOperation_VersionError)
    If mIndex > mEndIndex Then _
        Throw Error.InvalidOperation(InvalidOperation_EnumFinished)
    
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
