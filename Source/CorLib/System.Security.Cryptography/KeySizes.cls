VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeySizes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: KeySizes
'

''
' Represents a set of key sizes.
'
' @remarks All of the key sizes are represented by indicating a minimum key size and a maximum key size,
' then using a step to fill in key sizes between the two.
'
'@Folder("CorLib.System.Security.Cryptography")
Option Explicit
Implements IObject

Private mMinSize    As Long
Private mMaxSize    As Long
Private mSkipSize   As Long


''
' The mininum key size in bits.
'
' @return The mininum key size in bits.
'
Public Property Get MinSize() As Long
    MinSize = mMinSize
End Property

''
' The maximum key size in bits.
'
' @return The maximum key size in bits.
'
Public Property Get MaxSize() As Long
    MaxSize = mMaxSize
End Property

''
' The step size between keys in bits.
'
' @returns the step size in bits.'
'
Public Property Get SkipSize() As Long
    SkipSize = mSkipSize
End Property

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to test for equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

Friend Function IsLegalSize(ByVal BitLength As Long) As Boolean
    If mSkipSize = 0 Then
        IsLegalSize = (BitLength = mMinSize)
        Exit Function
    End If
    
    Dim LegalSize As Long
    For LegalSize = mMinSize To mMaxSize Step mSkipSize
        If BitLength = LegalSize Then
            IsLegalSize = True
            Exit Function
        End If
    Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MinSize As Long, ByVal MaxSize As Long, ByVal SkipSize As Long)
    mMinSize = MinSize
    mMaxSize = MaxSize
    mSkipSize = SkipSize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
