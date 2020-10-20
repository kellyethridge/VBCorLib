VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: CharEnumerator
'

''
' Used to enumerate the characters in a string from start to end.
'
' @remarks
' <p>Each character in the string is treated as an Integer value. This
' equates to using AscW for each character. There is no Unicode conversion
' using a Code Page.</p>
' <p>The return value is a 16-bit signed Integer. This means that
' characters above &h7fff (32767) with be negative. To change this to the
' corrisponding positive value, add &h10000 (65536) to the value.</p>
' @examples
' <PRE>
' Dim n As Variant
' Dim s As String
'
' s = "Hello"
'
'
' '' Iterate over a String using For..Each
' For Each n In NewCharEnumerator(s)
'     Debug.Print n
' Next n
' '' output
' '' 72
' '' 101
' '' 108
' '' 108
' '' 111
'
'
' '' Iterate over a String manually
' Dim e As CharEnumerator
' Set e = NewCharEnumerator(s)
'
'
' Do While e.MoveNext
'     Debug.Print e.Current
' Loop
' '' output
' '' 72
' '' 101
' '' 108
' '' 108
' '' 111
'</PRE></p>
'
' @see Constructors
' @see IEnumerator
'
'@Folder("CorLib.System")
Option Explicit
Implements IObject
Implements IEnumerator
Implements ICloneable

Private mBase       As EnumeratorBase
Private mString     As String
Private mChars()    As Integer
Private mCopyString As Boolean



''
' This allows the For..Each method to be used on this object.
'
' @return A For..Each compatible object used for enumeration.
' @remarks To use the CharEnumerator with the For..Each method,
' call the NewCharEnumerator method after the In keyword of
' the For..Each.<br>
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(Me)
End Function

''
' This returns the current item of the enumeration. The current
' item is called after a successful MoveNext.
'
' @return Returns the raw Unicode value for the current character in the string.
' @remarks The return value is a 16-bit signed Integer. This means that
' characters above &h7fff (32767) with be negative. To change this to the
' corrisponding positive value, add &h10000 (65536) to the value.
'
Public Property Get Current() As Integer
    Current = mChars(mBase.Index)
End Property

''
' Requests the enumerator moves to the next item being enumerated.
' Returns True on success, False otherwise. This is called prior
' to Current. If this is successful, Current will be called.
'
' @return Indication if the move was successful.
'
Public Function MoveNext() As Boolean
    MoveNext = mBase.MoveNext
End Function

''
' Requests the enumerator resets itself to begin enumerating from the beginning.
'
Public Sub Reset()
    mBase.Reset
End Sub

''
' Returns a clone of this instance of the enumerator.
'
' @return A clone of the CharEnumerator.
' @remarks <p>The cloned enumerator will begin enumerating at the location
' the original enumerator was at when the clone was created.</p>
' <p>The original string will always be copied when a clone is created.</p>
'
Public Function Clone() As CharEnumerator
    Set Clone = New CharEnumerator
    Clone.InitClone mBase.Clone, mString, mCopyString
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByVal CopyString As Boolean)
    Set mBase = Cor.NewEnumeratorBase(0, Len(s))
    InitChars s, CopyString
End Sub

Friend Sub InitClone(ByVal Base As EnumeratorBase, ByRef s As String, ByVal CopyString As Boolean)
    Set mBase = Base
    InitChars s, CopyString
End Sub

Private Sub Class_Terminate()
    FreeChars mChars
    
    If Not mCopyString Then
        StringPtr(mString) = vbNullPtr
    End If
End Sub

Private Sub InitChars(ByRef s As String, ByVal CopyString As Boolean)
    If CopyString Then
        mString = s
        mCopyString = True
    Else
'        mString = StringRef(s)
        StringPtr(mString) = StrPtr(s)
    End If

    mChars = AllocChars(mString)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    IEnumerator_Current = Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext
End Function

Private Sub IEnumerator_Reset()
    Reset
End Sub


