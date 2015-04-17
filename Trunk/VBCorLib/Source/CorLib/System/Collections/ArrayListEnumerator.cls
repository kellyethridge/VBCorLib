VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayListEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: ArrayListEnumerator
'

''
' This class is used to enumerate the elements of an ArrayList object.
'
Option Explicit
Implements IObject
Implements IEnumerator
Implements ICloneable

Private mList           As ArrayList
Private mStartIndex     As Long
Private mEndIndex       As Long
Private mIndex          As Long
Private mVersion        As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal List As ArrayList, ByVal Index As Long, ByVal Count As Long)
    Set mList = List
    mVersion = GetVersion(List)
    mStartIndex = Index
    mEndIndex = Index + Count - 1
    mIndex = Index - 1
End Sub

Friend Sub InitClone(ByVal List As ArrayList, ByVal StartIndex As Long, ByVal Index As Long)
    Init List, StartIndex, List.Count
    mIndex = Index
End Sub


' This is to allow code such as the following to function properly.
'
' For Each v In obj.GetEnumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Dim Ret As New ArrayListEnumerator
    Ret.InitClone mList, mStartIndex, mIndex
    Set ICloneable_Clone = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    If GetVersion(mList) <> mVersion Then _
        Error.InvalidOperation InvalidOperation_VersionError
    If mIndex < mStartIndex Then _
        Error.InvalidOperation InvalidOperation_EnumNotStarted
    If mIndex > mEndIndex Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
        
    Helper.MoveVariant IEnumerator_Current, mList(mIndex)
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If GetVersion(mList) <> mVersion Then _
        Error.InvalidOperation InvalidOperation_VersionError
    If mIndex > mEndIndex Then _
        Error.InvalidOperation InvalidOperation_EnumFinished
    
    If mIndex = mEndIndex Then
        mIndex = mEndIndex + 1
    Else
        mIndex = mIndex + 1
        IEnumerator_MoveNext = True
    End If
End Function

Private Sub IEnumerator_Reset()
    mIndex = mStartIndex - 1
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

