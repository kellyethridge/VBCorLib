VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecoderReplacementFallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: DecoderReplacementFallback
'

''
' Provides a failure-handling mechanism, called a fallback, for an encoded input byte sequence that cannot be converted to an output character. The fallback emits a user-specified replacement string instead of a decoded input byte sequence.
'
Option Explicit
Implements IObject
Implements DecoderFallback

Private Const DefaultStringProp As String = "DefaultString"

Private mDefaultString As String


''
' Creates a <a href="DecoderFallbackBuffer.html">DecoderFallbackBuffer</a> object that is initialized with the replacement string of this <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
'
' @return A <a href="DecoderFallbackBuffer.html">DecoderFallbackBuffer</a> object that specifies a string to use instead of the original decoding operation input.
Public Function CreateFallbackBuffer() As DecoderFallbackBuffer
    Set CreateFallbackBuffer = Cor.NewDecoderReplacementFallbackBuffer(Me)
End Function

''
' Gets the replacement string that is the value of the <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
'
' @return A substitute string that is emitted in place of an input byte sequence that cannot be decoded.
Public Property Get DefaultString() As String
    DefaultString = mDefaultString
End Property

''
' Gets the number of characters in the replacement string for the <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
'
' @return The number of characters in the string that is emitted in place of a byte sequence that cannot be decoded, that is, the length of the string returned by the <a href="DecoderReplacementFallback.get.DefaultString.html">DefaultString</a> property.
Public Property Get MaxCharCount() As Long
    MaxCharCount = Len(mDefaultString)
End Property

''
' Indicates whether the value of a specified object is equal to the <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
'
' @param Value A <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
' @return <b>True</b> if value is a <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object having a <a href="DecoderReplacementFallback.get.DefaultString.html">DefaultString</a> property that is equal to the <a href="DecoderReplacementFallback.get.DefaultString.html">DefaultString</a> property of the current <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object; otherwise, <b>False</b>.
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As DecoderReplacementFallback
    
    On Error GoTo Catch
    Set Other = Value
    Equals = Other.DefaultString = mDefaultString
    
Catch:
End Function

''
' Retrieves the hash code for the value of the <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
'
' @return The hash code of the value of the object.
' @remarks The value of a <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object is the value of its <a href="DecoderReplacementFallback.get.DefaultString.html">DefaultString</a> property.
Public Function GetHashCode() As Long
    GetHashCode = Object.GetHashCode(mDefaultString)
End Function

''
' Returns a string that represents the current object.
'
' @return A string that represents the current object.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Replacement As String)
    Dim FoundHigh   As Boolean
    Dim i           As Long
    
    For i = 0 To Len(Replacement) - 1
        If Char.IsSurrogateStr(Replacement, i) Then
            If Char.IsHighSurrogateStr(Replacement, i) Then
                If FoundHigh Then
                    Exit For
                End If
                
                FoundHigh = True
            End If
            
            If Char.IsLowSurrogateStr(Replacement, i) Then
                If Not FoundHigh Then
                    FoundHigh = True
                    Exit For
                End If
                
                FoundHigh = False
            End If
        ElseIf FoundHigh Then
            Exit For
        End If
    Next

    If FoundHigh Then _
        Error.Argument Argument_InvalidCharSequenceNoIndex, "Replacement"

    mDefaultString = Replacement
End Sub

Private Sub Class_Initialize()
    Init "?"
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mDefaultString = PropBag.ReadProperty(DefaultStringProp)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty DefaultStringProp, mDefaultString
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DecoderFallback
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DecoderFallback_CreateFallbackBuffer() As DecoderFallbackBuffer
    Set DecoderFallback_CreateFallbackBuffer = CreateFallbackBuffer
End Function

Private Function DecoderFallback_Equals(Value As Variant) As Boolean
    DecoderFallback_Equals = Equals(Value)
End Function

Private Function DecoderFallback_GetHashCode() As Long
    DecoderFallback_GetHashCode = GetHashCode
End Function

Private Property Get DecoderFallback_MaxCharCount() As Long
    DecoderFallback_MaxCharCount = MaxCharCount
End Property

Private Function DecoderFallback_ToString() As String
    DecoderFallback_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode()
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString()
End Function
