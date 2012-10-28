VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: StackEnumerator
'

''
' Class that is used to enumerate a Stack object.
'
Option Explicit
Implements IObject
Implements IEnumerator
Implements ICloneable

Private mStack      As Stack
Private mIndex      As Long
Private mItems()    As Variant
Private mResetPoint As Long
Private mVersion    As Long


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal st As Stack, ByVal ItemsSAPtr As Long)
    Set mStack = st
    SAPtr(mItems) = ItemsSAPtr
    mResetPoint = st.Count
    mVersion = st.Version
    mIndex = mResetPoint
End Sub

Friend Sub CloneHelper(ByVal st As Stack, ByVal ItemsSAPtr As Long, ByVal Index As Long)
    Call Init(st, ItemsSAPtr)
    mIndex = Index
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    SAPtr(mItems) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Object.Equals(Me, Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = ObjPtr(CUnk(Me))
End Function

Private Function IObject_ToString() As String
    IObject_ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim ret As New StackEnumerator
    Call ret.CloneHelper(mStack, SAPtr(mItems), mIndex)
    Set ICloneable_Clone = ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If mStack.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex = mResetPoint Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
        
    Call VariantCopy(IEnumerator_Current, mItems(mIndex))
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If mStack.Version <> mVersion Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_VersionError))
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))
        
    If mIndex = 0 Then
        mIndex = -1
    Else
        mIndex = mIndex - 1
        IEnumerator_MoveNext = True
    End If
End Function

Private Sub IEnumerator_Reset()
    mIndex = mResetPoint
End Sub
