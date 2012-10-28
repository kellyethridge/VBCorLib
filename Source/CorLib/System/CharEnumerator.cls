VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CharEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: CharEnumerator
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
' <h4>Example</h4>
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
Option Explicit
Implements IObject
Implements IEnumerator

Private mString     As WordBuffer
Private mIndex      As Long
Private mEndIndex   As Long
Private mStringCopy As String



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
    If mIndex < 0 Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumNotStarted))
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))

    Current = mString.Data(mIndex)
End Property

''
' Requests the enumerator moves to the next item being enumerated.
' Returns True on success, False otherwise. This is called prior
' to Current. If this is successful, Current will be called.
'
' @return Indication if the move was successful.
'
Public Function MoveNext() As Boolean
    If mIndex > mEndIndex Then _
        Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_EnumFinished))

    If mIndex = mEndIndex Then
        mIndex = mEndIndex + 1
    Else
        mIndex = mIndex + 1
        MoveNext = True
    End If
End Function

''
' Requests the enumerator resets itself to begin enumerating from the beginning.
'
Public Sub Reset()
    mIndex = -1
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByVal CopyString As Boolean)
    Dim Ptr As Long
    
    If CopyString Then
        mStringCopy = s
        Ptr = StrPtr(mStringCopy)
    Else
        Ptr = StrPtr(s)
    End If
    mEndIndex = Len(s) - 1
    Call InitWordBuffer(mString, Ptr, mEndIndex + 1)
    Reset
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mIndex = .ReadProperty("Index")
        mEndIndex = .ReadProperty("EndIndex")
        mStringCopy = .ReadProperty("String")
        Call InitWordBuffer(mString, StrPtr(mStringCopy), mEndIndex + 1)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Index", mIndex)
        Call .WriteProperty("EndIndex", mEndIndex)
        Call .WriteProperty("String", mStringCopy)
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
'   IEnumerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IEnumerator_Current() As Variant
    IEnumerator_Current = Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    IEnumerator_MoveNext = MoveNext
End Function

Private Sub IEnumerator_Reset()
    Call Reset
End Sub


