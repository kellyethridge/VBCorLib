VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Random"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: Random
'

''
' Provides a sequence of pseudo-random generated numbers.
'
' @see Constructors
'
Option Explicit
Implements IObject


''
' Returns the next generated number in the sequence.
'
' @param maxValue The value that the generated number must be less than.
' @return The next generated number in the sequence.
' @remarks The generated number will be equal or greater than zero and less than
' maxValue. The number will never equal maxValue.
Public Function NextLong(Optional ByVal MaxValue As Long = &H7FFFFFFF) As Long
    If MaxValue < 0 Then _
        Error.ArgumentOutOfRange "MaxValue", ArgumentOutOfRange_NeedNonNegNum
        
    NextLong = NextNumber(0, MaxValue)
End Function

''
' Returns the next generated number in the sequence within a specified range.
'
' @param minValue The lowerbound of the number to be generated.
' @param maxValue The upperbound of the number to be generated. Must be equal to or greater than minValue.
' @return A number generated within the specified range.
' @remarks The number generated will always be equal to or greater than minValue, and less than maxValue.
' The number will never equal maxValue.
Public Function NextRange(ByVal MinValue As Long, ByVal MaxValue As Long) As Long
    If MaxValue < MinValue Then _
        Throw Cor.NewArgumentOutOfRangeException("MinValue", Environment.GetResourceString(Argument_MinMaxValue, MinValue, MaxValue))
        
    NextRange = NextNumber(MinValue, MaxValue)
End Function

''
' Fills an array of bytes with random numbers.
'
' @param bytes The array to be filled.
Public Sub NextBytes(ByRef Bytes() As Byte)
    If SAPtr(Bytes) = vbNullPtr Then _
        Error.ArgumentNull "Bytes", ArgumentNull_Array
        
    Dim i As Long
    For i = LBound(Bytes) To UBound(Bytes)
        Bytes(i) = NextNumber(0, 256)    ' it won't reach 256, 255 is the max.
    Next i
End Sub

''
' Returns a random number between 0.0 and 1.0.
'
' @return The random number.
' @remarks The number will be equal to or greater than 0.0 and less than 1.0.
' The number will never equal 1.0.
Public Function NextDouble() As Double
    NextDouble = Rnd
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitRandom
End Sub

Friend Sub Init(ByVal Seed As Long)
    Rnd Seed
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NextNumber(ByVal MinValue As Long, ByVal MaxValue As Long) As Long
    NextNumber = Int(Rnd * (MaxValue - MinValue)) + MinValue
End Function


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
