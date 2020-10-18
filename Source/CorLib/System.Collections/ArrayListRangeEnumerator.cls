VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListRangeEnumerator"
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
'    Module: ArrayListRangeEnumerator
'
Option Explicit
Implements cObject
Implements IEnumerator

Private mList As ArrayList
Private mStartIndex As Long
Private mEndIndex As Long
Private mIndex As Long
Private mVersion As Long



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function



Friend Sub Init(ByVal list As ArrayList, ByVal index As Long, ByVal count As Long)
    Set mList = list
    mVersion = mList.Version
    mStartIndex = index
    mEndIndex = index + count - 1
    mIndex = index - 1
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException("ArrayList has changed.")
    If mIndex < mStartIndex Then _
        Throw Cor.NewInvalidOperationException("Enumeration has not started.")
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException("Enumeration has finished.")
    
    VariantCopyInd IEnumerator_Current, mList(mIndex)
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mList.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException("ArrayList has changed.")
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException("Enumeration has finished.")
        
    If mIndex < mEndIndex Then
        mIndex = mIndex + 1
        IEnumerator_MoveNext = True
    Else
        mIndex = mEndIndex + 1
    End If
End Function

Private Sub IEnumerator_Reset()
    mIndex = mStartIndex - 1
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   cObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function cObject_Equals(value As Variant) As Boolean
    If IsObject(value) Then
        cObject_Equals = (value Is Me)
    End If
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = ObjPtr(Me)
End Function

Private Function cObject_ToString() As String
    cObject_ToString = App.Title & "." & TypeName(Me)
End Function


