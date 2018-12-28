Attribute VB_Name = "Mathematics"
'The MIT License (MIT)
'Copyright (c) 2018 Kelly Ethridge
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
' Module: Mathematics
'
Option Explicit

Public Function ShiftRightInt32(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Jost Schwider, jost@schwider.de, 20010928
    If ShiftCount = 0 Then
        ShiftRightInt32 = Value
    Else
        If Value And &H80000000 Then
            Value = (Value And &H7FFFFFFF) \ 2 Or &H40000000
            ShiftCount = ShiftCount - 1
        End If
        
        Select Case ShiftCount
            Case 0&:  ShiftRightInt32 = Value
            Case 1&:  ShiftRightInt32 = Value \ &H2&
            Case 2&:  ShiftRightInt32 = Value \ &H4&
            Case 3&:  ShiftRightInt32 = Value \ &H8&
            Case 4&:  ShiftRightInt32 = Value \ &H10&
            Case 5&:  ShiftRightInt32 = Value \ &H20&
            Case 6&:  ShiftRightInt32 = Value \ &H40&
            Case 7&:  ShiftRightInt32 = Value \ &H80&
            Case 8&:  ShiftRightInt32 = Value \ &H100&
            Case 9&:  ShiftRightInt32 = Value \ &H200&
            Case 10&: ShiftRightInt32 = Value \ &H400&
            Case 11&: ShiftRightInt32 = Value \ &H800&
            Case 12&: ShiftRightInt32 = Value \ &H1000&
            Case 13&: ShiftRightInt32 = Value \ &H2000&
            Case 14&: ShiftRightInt32 = Value \ &H4000&
            Case 15&: ShiftRightInt32 = Value \ &H8000&
            Case 16&: ShiftRightInt32 = Value \ &H10000
            Case 17&: ShiftRightInt32 = Value \ &H20000
            Case 18&: ShiftRightInt32 = Value \ &H40000
            Case 19&: ShiftRightInt32 = Value \ &H80000
            Case 20&: ShiftRightInt32 = Value \ &H100000
            Case 21&: ShiftRightInt32 = Value \ &H200000
            Case 22&: ShiftRightInt32 = Value \ &H400000
            Case 23&: ShiftRightInt32 = Value \ &H800000
            Case 24&: ShiftRightInt32 = Value \ &H1000000
            Case 25&: ShiftRightInt32 = Value \ &H2000000
            Case 26&: ShiftRightInt32 = Value \ &H4000000
            Case 27&: ShiftRightInt32 = Value \ &H8000000
            Case 28&: ShiftRightInt32 = Value \ &H10000000
            Case 29&: ShiftRightInt32 = Value \ &H20000000
            Case 30&: ShiftRightInt32 = Value \ &H40000000
            Case 31&: ShiftRightInt32 = &H0&
        End Select
    End If
End Function

Public Function ShiftRightInt64(ByRef Value As DLong, ByVal ShiftCount As Long) As DLong
    Dim BitsToMove As Long
    
    If ShiftCount < 64 Then
        If ShiftCount < 32 Then
            ShiftRightInt64.LoDWord = ShiftRightInt32(Value.LoDWord, ShiftCount)
            BitsToMove = Value.HiDWord And (Powers(ShiftCount) - 1)
            ShiftRightInt64.LoDWord = ShiftRightInt64.LoDWord Or ShiftLeftInt32(BitsToMove, 32 - ShiftCount)
            ShiftRightInt64.HiDWord = ShiftRightInt32(Value.HiDWord, ShiftCount)
        Else
            ShiftRightInt64.LoDWord = ShiftRightInt32(Value.HiDWord, ShiftCount - 32)
        End If
    End If
End Function

Public Function ShiftLeftInt32(ByVal Value As Long, ByVal ShiftCount As Long) As Long
' by Jost Schwider, jost@schwider.de, 20011001
    Select Case ShiftCount
        Case 0&
            ShiftLeftInt32 = Value
        Case 1&
            If Value And &H40000000 Then
              ShiftLeftInt32 = (Value And &H3FFFFFFF) * &H2& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3FFFFFFF) * &H2&
            End If
        Case 2&
            If Value And &H20000000 Then
              ShiftLeftInt32 = (Value And &H1FFFFFFF) * &H4& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1FFFFFFF) * &H4&
            End If
        Case 3&
            If Value And &H10000000 Then
              ShiftLeftInt32 = (Value And &HFFFFFFF) * &H8& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HFFFFFFF) * &H8&
            End If
        Case 4&
            If Value And &H8000000 Then
              ShiftLeftInt32 = (Value And &H7FFFFFF) * &H10& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7FFFFFF) * &H10&
            End If
        Case 5&
            If Value And &H4000000 Then
              ShiftLeftInt32 = (Value And &H3FFFFFF) * &H20& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3FFFFFF) * &H20&
            End If
        Case 6&
            If Value And &H2000000 Then
              ShiftLeftInt32 = (Value And &H1FFFFFF) * &H40& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1FFFFFF) * &H40&
            End If
        Case 7&
            If Value And &H1000000 Then
              ShiftLeftInt32 = (Value And &HFFFFFF) * &H80& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HFFFFFF) * &H80&
            End If
        Case 8&
            If Value And &H800000 Then
              ShiftLeftInt32 = (Value And &H7FFFFF) * &H100& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7FFFFF) * &H100&
            End If
        Case 9&
            If Value And &H400000 Then
              ShiftLeftInt32 = (Value And &H3FFFFF) * &H200& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3FFFFF) * &H200&
            End If
        Case 10&
            If Value And &H200000 Then
              ShiftLeftInt32 = (Value And &H1FFFFF) * &H400& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1FFFFF) * &H400&
            End If
        Case 11&
            If Value And &H100000 Then
              ShiftLeftInt32 = (Value And &HFFFFF) * &H800& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HFFFFF) * &H800&
            End If
        Case 12&
            If Value And &H80000 Then
              ShiftLeftInt32 = (Value And &H7FFFF) * &H1000& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7FFFF) * &H1000&
            End If
        Case 13&
            If Value And &H40000 Then
              ShiftLeftInt32 = (Value And &H3FFFF) * &H2000& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3FFFF) * &H2000&
            End If
        Case 14&
            If Value And &H20000 Then
              ShiftLeftInt32 = (Value And &H1FFFF) * &H4000& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1FFFF) * &H4000&
            End If
        Case 15&
            If Value And &H10000 Then
              ShiftLeftInt32 = (Value And &HFFFF&) * &H8000& Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HFFFF&) * &H8000&
            End If
        Case 16&
            If Value And &H8000& Then
              ShiftLeftInt32 = (Value And &H7FFF&) * &H10000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7FFF&) * &H10000
            End If
        Case 17&
            If Value And &H4000& Then
              ShiftLeftInt32 = (Value And &H3FFF&) * &H20000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3FFF&) * &H20000
            End If
        Case 18&
            If Value And &H2000& Then
              ShiftLeftInt32 = (Value And &H1FFF&) * &H40000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1FFF&) * &H40000
            End If
        Case 19&
            If Value And &H1000& Then
              ShiftLeftInt32 = (Value And &HFFF&) * &H80000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HFFF&) * &H80000
            End If
        Case 20&
            If Value And &H800& Then
              ShiftLeftInt32 = (Value And &H7FF&) * &H100000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7FF&) * &H100000
            End If
        Case 21&
            If Value And &H400& Then
              ShiftLeftInt32 = (Value And &H3FF&) * &H200000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3FF&) * &H200000
            End If
        Case 22&
            If Value And &H200& Then
              ShiftLeftInt32 = (Value And &H1FF&) * &H400000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1FF&) * &H400000
            End If
        Case 23&
            If Value And &H100& Then
              ShiftLeftInt32 = (Value And &HFF&) * &H800000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HFF&) * &H800000
            End If
        Case 24&
            If Value And &H80& Then
              ShiftLeftInt32 = (Value And &H7F&) * &H1000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7F&) * &H1000000
            End If
        Case 25&
            If Value And &H40& Then
              ShiftLeftInt32 = (Value And &H3F&) * &H2000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3F&) * &H2000000
            End If
        Case 26&
            If Value And &H20& Then
              ShiftLeftInt32 = (Value And &H1F&) * &H4000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1F&) * &H4000000
            End If
        Case 27&
            If Value And &H10& Then
              ShiftLeftInt32 = (Value And &HF&) * &H8000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &HF&) * &H8000000
            End If
        Case 28&
            If Value And &H8& Then
              ShiftLeftInt32 = (Value And &H7&) * &H10000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H7&) * &H10000000
            End If
        Case 29&
            If Value And &H4& Then
              ShiftLeftInt32 = (Value And &H3&) * &H20000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H3&) * &H20000000
            End If
        Case 30&
            If Value And &H2& Then
              ShiftLeftInt32 = (Value And &H1&) * &H40000000 Or &H80000000
            Else
              ShiftLeftInt32 = (Value And &H1&) * &H40000000
            End If
        Case 31&
            If Value And &H1& Then
              ShiftLeftInt32 = &H80000000
            Else
              ShiftLeftInt32 = &H0&
            End If
    End Select
End Function

Public Function RRotate(ByVal Value As Long, ByVal Count As Long) As Long
    RRotate = Helper.ShiftRight(Value, Count) Or Helper.ShiftLeft(Value, 32 - Count)
End Function

Public Function LRotate(ByVal Value As Long, ByVal Count As Long) As Long
    LRotate = Helper.ShiftLeft(Value, Count) Or Helper.ShiftRight(Value, 32 - Count)
End Function

''
' Modulus method used for large values held within currency datatypes.
'
' @param x The value to be divided.
' @param y The value used to divide.
' @return The remainder of the division.
'
Public Function Modulus(ByVal x As Currency, ByVal y As Currency) As Currency
  Modulus = x - (y * Fix(x / y))
End Function

Public Function SwapEndian(ByVal Value As Long) As Long
    SwapEndian = (((Value And &HFF000000) \ &H1000000) And &HFF&) Or _
                 ((Value And &HFF0000) \ &H100&) Or _
                 ((Value And &HFF00&) * &H100&) Or _
                 ((Value And &H7F&) * &H1000000)
    If (Value And &H80&) Then SwapEndian = SwapEndian Or &H80000000
End Function

