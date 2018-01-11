Attribute VB_Name = "BigNumberMath"
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
' Module: Mathematics
'

''
' This module contains the primary arithmetic algorithms used by the library.
'
' There are two sets of each function in this module. The standard functions are
' optimized and should only be run when compiled (with Integer Overflow turned off.)
'
' The second set of functions are to provide safe versions that can execute within an IDE environment.
'
Option Explicit

''
' This contains all the information about a number. The information can be easily
' passed around as a group instead of trying to pass individual parameters.
'
Public Type BigNumber
    Digits()    As Integer
    Precision   As Long
    Sign        As Long
End Type

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

Public Function Equals(ByRef x As BigNumber, ByRef y As BigNumber) As Boolean
    If x.Sign <> y.Sign Then
        Exit Function
    ElseIf x.Precision <> y.Precision Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To x.Precision - 1
        If x.Digits(i) <> y.Digits(i) Then
            Exit Function
        End If
    Next
    
    Equals = True
End Function

#If Release Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' These Release methods must be compiled with Interger Overflow
' checks off. The methods must also pass all tests once compiled.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' This is the basic implementation of a gradeschool style
' addition of two n-place numbers.
'
' Ref: The Art of Computer Programming 4.3.1.A
'
Public Function GradeSchoolAdd(ByRef u As BigNumber, ByRef v As BigNumber) As Integer()
    Dim uExtDigit   As Long
    Dim vExtDigit   As Long
    Dim sum()       As Integer

    If u.Sign = Negative Then
        uExtDigit = &HFFFF&
    End If
    
    If v.Sign = Negative Then
        vExtDigit = &HFFFF&
    End If

    If u.Precision >= v.Precision Then
        ReDim sum(0 To u.Precision)
    Else
        ReDim sum(0 To v.Precision)
    End If

    Dim i As Long
    Dim k As Long
    For i = 0 To UBound(sum)
        Dim uDigit  As Long
        Dim vDigit  As Long
        
        If i < u.Precision Then
            uDigit = u.Digits(i) And &HFFFF&
        Else
            uDigit = uExtDigit
        End If
        
        If i < v.Precision Then
            vDigit = v.Digits(i) And &HFFFF&
        Else
            vDigit = vExtDigit
        End If
        
        k = uDigit + vDigit + k ' this is the only change from subtraction
        sum(i) = k And &HFFFF&
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolAdd = sum
End Function

''
' This is the basic implementation of a gradeschool style
' subtraction of two n-place numbers.
'
' Ref: The Art of Computer Programming 4.3.1.S
'
Public Function GradeSchoolSubtract(ByRef u As BigNumber, ByRef v As BigNumber) As Integer()
    Dim uExtDigit   As Long
    Dim vExtDigit    As Long
    Dim Difference() As Integer

    If u.Sign = Negative Then
        uExtDigit = &HFFFF&
    End If
    
    If v.Sign = Negative Then
        vExtDigit = &HFFFF&
    End If

    If u.Precision >= v.Precision Then
        ReDim Difference(0 To u.Precision)
    Else
        ReDim Difference(0 To v.Precision)
    End If

    Dim i       As Long
    Dim k       As Long
    For i = 0 To UBound(Difference)
        Dim uDigit  As Long
        Dim vDigit  As Long
        
        If i < u.Precision Then
            uDigit = u.Digits(i) And &HFFFF&
        Else
            uDigit = uExtDigit
        End If
        
        If i < v.Precision Then
            vDigit = v.Digits(i) And &HFFFF&
        Else
            vDigit = vExtDigit
        End If
        
        k = uDigit - vDigit + k ' this is the only change from addition
        Difference(i) = k And &HFFFF&
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolSubtract = Difference
End Function

''
' This is a straight forward implementation of Knuth's algorithm.
'
' Ref: The Art of Computer Programming 4.3.1.M
'
Public Function GradeSchoolMultiply(ByRef u As BigNumber, ByRef v As BigNumber) As Integer()
    Dim Product()   As Integer
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    
    ReDim Product(0 To u.Precision + v.Precision)
    
    For j = 0 To v.Precision - 1
        Dim d As Long
        
        d = v.Digits(j) And &HFFFF&
        k = 0
        
        For i = 0 To u.Precision - 1
            k = d * (u.Digits(i) And &HFFFF&) + (Product(i + j) And &HFFFF&) + k
            Product(i + j) = k And &HFFFF&
            k = ((k And &HFFFF0000) \ &H10000) And &HFFFF&
        Next i
        
        Product(i + j) = k And &HFFFF&
    Next j
    
    GradeSchoolMultiply = Product
End Function

''
' This is an implementation of Knuth's algorithm.
'
' As simple as division would seem to be in the real world, implementing it at such
' a low level has its own sets of problems. After careful study of Knuth's algorithm
' I finally came up with the following implmentation. The steps in the book are
' marked inline with the code as close as possible.
'
' Ref: The Art of Computer Programming 4.3.1.D
'
Public Function GradeSchoolDivide(ByRef u As BigNumber, ByRef v As BigNumber, ByRef remainder() As Integer, ByVal IncludeRemainder As Boolean) As Integer()
    Dim n As Long
    Dim m As Long
    Dim d As Long
    
    n = v.Precision
    m = u.Precision - n
      
    ' test if the divisor is shorter than the dividend, if so then just
    ' return a 0 quotient and the dividend as the remainder, if needed.
    If m < 0 Then
        If IncludeRemainder Then
            ReDim remainder(0 To u.Precision)
            CopyMemory remainder(0), u.Digits(0), u.Precision * 2
        End If
        
        GradeSchoolDivide = Cor.NewIntegers()
        Exit Function
    End If
    
    ' ** D1 Start **
    If (u.Precision - 1) = UBound(u.Digits) Then
        ReDim Preserve u.Digits(0 To u.Precision)
    End If
    
    u.Digits(u.Precision) = 0
    u.Precision = u.Precision + 1

    d = &H10000 \ (1 + (v.Digits(n - 1) And &HFFFF&))
    
    If d > 1 Then
        SingleInPlaceMultiply u, d
        SingleInPlaceMultiply v, d
    End If
    ' ** D1 End **
    
    Dim Quotient() As Integer
    ReDim Quotient(0 To m + 1)
    
    Dim vDigit  As Long
    Dim vDigit2 As Long
    
    ' this is the Vn-1 digit used repeatedly in step D3.
    vDigit = v.Digits(n - 1) And &HFFFF&
    
    ' this is the Vn-2 digit used repeatedly in step D3.
    If n - 2 >= 0 Then
        vDigit2 = v.Digits(n - 2) And &HFFFF&
    End If
    
    Dim qXu() As Integer    ' cache the array to prevent constant allocate/deallocate
    ReDim qXu(0 To n)       ' the array will be reused for multiplication
    
    ' this is an optimistic caching to be used incase
    ' a negative value is encountered. the same value
    ' will always be used regardless, so cache it here.
    Dim q2 As Long
    Dim r2 As Long
    q2 = &H7FFFFFFF \ vDigit
    r2 = &H7FFFFFFF - (q2 * vDigit) + 1
    
    Dim j       As Long
    Dim rHat    As Long
    Dim qHat    As Long
    
    ' ** D2 Start **
    For j = m To 0 Step -1
        Dim WordU As Long
        
        ' ** D3 Start **
        ' since we are shifting left, it is possible that we could turn wordu
        ' into a negative value and will need to deal with it differently later on.
        WordU = ((u.Digits(j + n) And &HFFFF&) * &H10000) Or (u.Digits(j + n - 1) And &HFFFF&)
        
        ' We have to deal with dividing negatives. They need to work like unsigned.
        If WordU And &H80000000 Then
            Dim q1 As Long
            q1 = (WordU And &H7FFFFFFF) \ vDigit
            rHat = (WordU And &H7FFFFFFF) - (q1 * vDigit) + r2
            
            If rHat >= vDigit Then
                q1 = q1 + 1
                rHat = rHat - vDigit
            End If

            qHat = q1 + q2
        Else
            qHat = WordU \ vDigit
            rHat = WordU - qHat * vDigit
        End If
        
        Do
            If qHat < &H10000 Then
                Dim qHatDigits As Long
                Dim rHatDigits As Long

                qHatDigits = (qHat * (v.Digits(n - 2) And &HFFFF&))
                rHatDigits = (rHat * &H10000) + (u.Digits(j + n - 2) And &HFFFF&)
                
                If (qHatDigits - &H80000000) <= (rHatDigits - &H80000000) Then
                    Exit Do
                End If
            End If
            
            qHat = qHat - 1
            rHat = rHat + vDigit
        Loop While rHat < &H10000
        ' ** D3 End **
        
        ' ** D4 Start **
        SinglePlaceMultiply v.Digits, n, qHat, qXu
        
        Dim Borrowed As Boolean
        Borrowed = MultiInPlaceSubtract(u.Digits, j, qXu)
        ' ** D4 End **
        
        ' ** D5 Start **
        If Borrowed Then
            ' ** D6 Start **
            qHat = qHat - 1
            MultiInPlaceAdd u.Digits, j, v.Digits
            ' ** D6 End **
        End If
        ' ** D5 End **
        
        Quotient(j) = qHat And &HFFFF&
    Next j
    ' ** D2 End **
    
    ' ** D8 Start **
    If IncludeRemainder Then
        If d > 1 Then
            remainder = SinglePlaceDivide(u.Digits, n, d)
        Else
            remainder = u.Digits
        End If
    End If
    ' ** D8 End **
    
    GradeSchoolDivide = Quotient
End Function

''
' Performs a single in-place division by 10, returning the remainder.
'
' The buffer is modified by this routine.
'
Public Function SingleInPlaceDivideBy10(ByRef n As BigNumber) As Long
    Dim R As Long
    Dim i As Long
    Dim f As Boolean
    Dim d As Long

    For i = n.Precision - 1 To 0 Step -1
        R = (R * &H10000) + (n.Digits(i) And &HFFFF&)
        d = R \ 10
        n.Digits(i) = d And &HFFFF&
        R = R - (d * 10)

        If Not f Then
            If n.Digits(i) = 0 Then
                n.Precision = n.Precision - 1
            Else
                f = True
            End If
        End If
    Next i

    SingleInPlaceDivideBy10 = R
End Function

''
' Performs a Two's Complement on the number, effectively negating it.
'
' The number buffer is modified by this routine. It will also reallocate
' the buffer if necessary.
'
Public Sub Negate(ByRef n As BigNumber)
    ' this is to handle situations like FFFF => FFFF0001.
    If n.Sign = Positive Then
        If n.Digits(n.Precision - 1) And &H8000 Then
            If n.Precision > UBound(n.Digits) Then
                ReDim Preserve n.Digits(0 To n.Precision)
            End If
            
            n.Digits(n.Precision) = 0
            n.Precision = n.Precision + 1
        End If
    End If

    Dim k As Long
    Dim i As Long
    
    k = 1
    
    For i = 0 To n.Precision - 1
        k = k + ((n.Digits(i) Xor &HFFFF) And &HFFFF&)
        n.Digits(i) = k And &HFFFF&
        k = (k And &HFFFF0000) \ &H10000
    Next i

    n.Sign = 0 - n.Sign
End Sub

''
' Performs a single in-place multiplication within the original array.
'
' The number buffer is modified by this routine. It is assumed the buffer
' is large enough to handle the larger result.
'
Public Sub SingleInPlaceMultiply(ByRef n As BigNumber, ByVal Value As Long)
    Dim Result  As Long
    Dim i       As Long

    For i = 0 To n.Precision - 1
        Result = Result + Value * (n.Digits(i) And &HFFFF&)
        n.Digits(i) = Result And &HFFFF&
        Result = ((Result And &HFFFF0000) \ &H10000) And &HFFFF&
    Next i

    If Result > 0 Then
        n.Precision = n.Precision + 1
        n.Digits(i) = Result And &HFFFF&
    End If
End Sub

''
' Performs a single in-place addition within the original array.
'
' The number buffer must be largest enough to handle any overflow.
'
Public Sub SingleInPlaceAdd(ByRef n As BigNumber, ByVal Value As Long)
    Dim i As Long
    
    Do While Value > 0
        If i >= n.Precision Then
            n.Precision = n.Precision + 1
        End If
        
        Value = Value + (n.Digits(i) And &HFFFF&)
        n.Digits(i) = Value And &HFFFF&
        Value = ((Value And &HFFFF0000) \ &H10000) And &HFFFF&
        i = i + 1
    Loop
End Sub

''
' This is a support routine for division.
'
Private Sub SinglePlaceMultiply(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, ByRef w() As Integer)
    Dim k As Long
    Dim i As Long
    
    For i = 0 To Length - 1
        k = k + (v * (u(i) And &HFFFF&))
        w(i) = k And &HFFFF&
        k = ((k And &HFFFF0000) \ &H10000) And &HFFFF&
    Next

    w(Length) = k And &HFFFF&
End Sub

''
' This is a support routine for division.
'
Private Function MultiInPlaceSubtract(ByRef u() As Integer, ByVal StartIndex As Long, ByRef v() As Integer) As Boolean
    Dim k       As Long
    Dim Result  As Long
    Dim d       As Long
    Dim i       As Long
    Dim j       As Long
    Dim ubv     As Long
    
    ubv = UBound(v)
    
    For i = StartIndex To UBound(u)
        If j <= ubv Then
            d = v(j) And &HFFFF&
        Else
            d = 0
        End If
        
        Result = Result + ((u(i) And &HFFFF&) - d) + k
        
        If Result < 0 Then
            Result = Result + &H10000
            k = -1
        Else
            k = 0
        End If
        
        u(i) = Result And &HFFFF&
        Result = ((Result And &HFFFF0000) \ &H10000) And &HFFFF&
        j = j + 1
    Next i
    
    MultiInPlaceSubtract = k
End Function

''
' Performs an addition between two arrays, placing the result in the first array.
'
Private Sub MultiInPlaceAdd(ByRef u() As Integer, ByVal StartIndex As Long, ByRef v() As Integer)
    Dim Result  As Long
    Dim i       As Long
    Dim j       As Long
    Dim d       As Long
    Dim ubv     As Long
    ubv = UBound(v)
    
    For i = StartIndex To UBound(u)
        If j <= ubv Then
            d = v(j) And &HFFFF&
        Else
            d = 0
        End If
        
        Result = Result + (u(i) And &HFFFF&) + d
        u(i) = Result And &HFFFF&
        Result = ((Result And &HFFFF0000) \ &H10000) And &HFFFF&
        j = j + 1
    Next i
End Sub

''
' Divides an array by a single digit (16bit) value, returning the quotient and remainder.
'
Public Function SinglePlaceDivide(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, Optional ByRef remainder As Long) As Integer()
    Dim R   As Long
    Dim q() As Integer
    Dim q2  As Long
    Dim r2  As Long
        
    ReDim q(0 To Length)
    v = v And &HFFFF&
    q2 = &H7FFFFFFF \ v
    r2 = &H7FFFFFFF - (q2 * v) + 1
    
    Dim i As Long
    For i = Length - 1 To 0 Step -1
        R = R * &H10000 + (u(i) And &HFFFF&)
        
        If R And &H80000000 Then
            Dim q1 As Long
            q1 = (R And &H7FFFFFFF) \ v
            R = (R And &H7FFFFFFF) - (q1 * v) + r2

            If R >= v Then
                q1 = q1 + 1
                R = R - v
            End If

            q(i) = q1 + q2
        Else
            q(i) = R \ v
            R = R - (q(i) And &HFFFF&) * v
        End If
    Next
    
    remainder = R
    SinglePlaceDivide = q
End Function

Public Sub ApplyTwosComplement(ByRef n() As Integer)
    Dim c As Long
    Dim i As Long
    
    c = 1
    
    For i = 0 To UBound(n)
        c = ((n(i) Xor &HFFFF) And &HFFFF&) + c
        n(i) = c And &HFFFF&
        c = (c And &HFFFF0000) \ &H10000
    Next
End Sub

#Else
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' These are IDE safe versions of the math routines.
'
' The routines are not optimized, they are provided only to
' allow this library to function safely while in development.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ApplyTwosComplement(ByRef n() As Integer)
    Dim c As Long
    Dim i As Long
    
    c = 1
    
    For i = 0 To UBound(n)
        c = ((n(i) Xor &HFFFF) And &HFFFF&) + c
        n(i) = AsWord(c)
        c = RightShift16(c)
    Next
End Sub

Public Function GradeSchoolMultiply(ByRef u As BigNumber, ByRef v As BigNumber) As Integer()
    Dim Product()   As Integer
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
        
    ReDim Product(0 To u.Precision + v.Precision)
    
    For i = 0 To v.Precision - 1
        k = 0
        
        For j = 0 To u.Precision - 1
            k = UInt16x16To32(v.Digits(i), u.Digits(j)) + (Product(i + j) And &HFFFF&) + k
            Product(i + j) = AsWord(k)
            k = RightShift16(k)
        Next j
        
        Product(i + j) = AsWord(k)
    Next
    
    GradeSchoolMultiply = Product
End Function

Public Sub SingleInPlaceMultiply(ByRef n As BigNumber, ByVal Value As Long)
    Dim k As Long
    Dim i As Long
    
    For i = 0 To n.Precision - 1
        k = UInt16x16To32(n.Digits(i), Value) + k
        n.Digits(i) = AsWord(k)
        k = RightShift16(k)
    Next
    
    If k Then
        n.Digits(n.Precision) = AsWord(k)
        n.Precision = n.Precision + 1
    End If
End Sub

Public Sub SingleInPlaceAdd(ByRef n As BigNumber, ByVal Value As Integer)
    Dim k As Long
    Dim i As Long
    
    k = Value And &HFFFF&
    
    Do While k > 0
        If i >= n.Precision Then
            n.Precision = n.Precision + 1
        End If
        
        k = (n.Digits(i) And &HFFFF&) + k
        n.Digits(i) = AsWord(k)
        k = RightShift16(k)
        i = i + 1
    Loop
End Sub

Public Sub Negate(ByRef n As BigNumber)
    Dim k As Long
    Dim i As Long

    k = 1
 
    ' this is to handle situations like FFFF => FFFF0001.
    If n.Sign = Sign.Positive Then
        If n.Digits(n.Precision - 1) And &H8000 Then
            If n.Precision > UBound(n.Digits) Then
                ReDim Preserve n.Digits(0 To n.Precision)
            End If

            n.Digits(n.Precision) = 0
            n.Precision = n.Precision + 1
        End If
    ElseIf n.Sign = Negative Then
        If n.Precision > 1 And n.Digits(n.Precision - 1) = &HFFFF Then
            n.Precision = n.Precision - 1
            n.Digits(n.Precision) = 0
        End If
    End If

    For i = 0 To n.Precision - 1
        k = k + ((n.Digits(i) Xor &HFFFF) And &HFFFF&)
        n.Digits(i) = AsWord(k)
        k = RightShift16(k)
    Next

    n.Sign = 0 - n.Sign
End Sub

Public Function SingleInPlaceDivideBy10(ByRef n As BigNumber) As Long
    Dim R As Long
    Dim i As Long
    Dim f As Boolean
    Dim d As Long

    For i = n.Precision - 1 To 0 Step -1
        R = (R * &H10000) + (n.Digits(i) And &HFFFF&)
        d = R \ 10
        n.Digits(i) = AsWord(d)
        R = R - (d * 10)

        If Not f Then
            If n.Digits(i) = 0 Then
                n.Precision = n.Precision - 1
            Else
                f = True
            End If
        End If
    Next

    SingleInPlaceDivideBy10 = R
End Function

Public Function GradeSchoolAdd(ByRef u As BigNumber, ByRef v As BigNumber) As Integer()
    Dim uExtDigit   As Long
    Dim vExtDigit   As Long
    Dim sum()       As Integer

    If u.Sign = Negative Then
        uExtDigit = &HFFFF&
    End If
    
    If v.Sign = Negative Then
        vExtDigit = &HFFFF&
    End If

    If u.Precision >= v.Precision Then
        ReDim sum(0 To u.Precision)
    Else
        ReDim sum(0 To v.Precision)
    End If
    
    Dim i As Long
    Dim k As Long
    For i = 0 To UBound(sum)
        Dim uDigit  As Long
        Dim vDigit  As Long
        
        If i < u.Precision Then
            uDigit = u.Digits(i) And &HFFFF&
        Else
            uDigit = uExtDigit
        End If
        
        If i < v.Precision Then
            vDigit = v.Digits(i) And &HFFFF&
        Else
            vDigit = vExtDigit
        End If
        
        k = uDigit + vDigit + k ' this is the only change for the subtraction
        sum(i) = AsWord(k)
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolAdd = sum
End Function

Public Function GradeSchoolSubtract(ByRef u As BigNumber, ByRef v As BigNumber) As Integer()
    Dim uExtDigit       As Long
    Dim vExtDigit       As Long
    Dim Difference()    As Integer

    If u.Sign = Negative Then
        uExtDigit = &HFFFF&
    End If
    
    If v.Sign = Negative Then
        vExtDigit = &HFFFF&
    End If
    
    If u.Precision >= v.Precision Then
        ReDim Difference(0 To u.Precision)
    Else
        ReDim Difference(0 To v.Precision)
    End If
    
    Dim i As Long
    Dim k As Long
    For i = 0 To UBound(Difference)
        Dim uDigit  As Long
        Dim vDigit  As Long
        
        If i < u.Precision Then
            uDigit = u.Digits(i) And &HFFFF&
        Else
            uDigit = uExtDigit
        End If
        
        If i < v.Precision Then
            vDigit = v.Digits(i) And &HFFFF&
        Else
            vDigit = vExtDigit
        End If

        k = uDigit - vDigit + k ' this is the only change for the addition
        Difference(i) = AsWord(k)
        k = (k And &HFFFF0000) \ &H10000
    Next i
    
    GradeSchoolSubtract = Difference
End Function

Public Function GradeSchoolDivide(ByRef u As BigNumber, ByRef v As BigNumber, ByRef remainder() As Integer, ByVal IncludeRemainder As Boolean) As Integer()
    Dim n As Long
    Dim m As Long
    Dim d As Long
    
    n = v.Precision
    m = u.Precision - n
     
    If m < 0 Then
        If IncludeRemainder Then
            ReDim remainder(0 To u.Precision)
            CopyMemory remainder(0), u.Digits(0), u.Precision * 2
        End If
        
        GradeSchoolDivide = Cor.NewIntegers()
        Exit Function
    End If
    
    If (u.Precision - 1) = UBound(u.Digits) Then
        ReDim Preserve u.Digits(0 To u.Precision)
    End If
    
    u.Digits(u.Precision) = 0
    u.Precision = u.Precision + 1
        
    d = &H10000 \ (1 + (v.Digits(n - 1) And &HFFFF&))
    
    If d > 1 Then
        SingleInPlaceMultiply u, d
        SingleInPlaceMultiply v, d
    End If
    
    Dim Quotient() As Integer
    ReDim Quotient(0 To m + 1)
    
    Dim vDigit As Integer
    vDigit = v.Digits(n - 1)
    
    Dim vDigit2 As Long
    If n - 2 >= 0 Then
        vDigit2 = v.Digits(n - 2) And &HFFFF&
    End If
    
    Dim qTimesu() As Integer
    ReDim qTimesu(0 To n)
    
    Dim j       As Long
    Dim rHat    As Long
    Dim qHat    As Long
    For j = m To 0 Step -1
        Dim WordU As Long
        
        WordU = Make32(u.Digits(j + n), u.Digits(j + n - 1))
        qHat = UInt32d16To32(WordU, vDigit)
        rHat = UInt32m16To32(WordU, vDigit)
        
        Do
            If qHat < &H10000 Then
                If UInt32Compare(UInt32x16To32(qHat, v.Digits(n - 2)), Helper.ShiftLeft(rHat, 16) + (u.Digits(j + n - 2) And &HFFFF&)) <= 0 Then
                    Exit Do
                End If
            End If
            
            qHat = qHat - 1
            rHat = rHat + (vDigit And &HFFFF&)
        Loop While rHat < &H10000
        
        SinglePlaceMultiply v.Digits, n, qHat, qTimesu
        
        Dim borrow As Boolean
        borrow = MultiInPlaceSubtract(u.Digits, j, qTimesu)
        
        If borrow Then
            qHat = qHat - 1
            MultiInPlaceAdd u.Digits, j, v.Digits
        End If
        
        Quotient(j) = AsWord(qHat)
    Next
    
    If IncludeRemainder Then
        If d > 1 Then
            remainder = SinglePlaceDivide(u.Digits, n, d)
        Else
            remainder = u.Digits
        End If
    End If
    
    GradeSchoolDivide = Quotient
End Function

Private Function UInt32x16To32(ByVal x As Long, ByVal y As Integer) As Long
    Dim v As Currency
    Dim w As Currency
    
    v = y And &HFFFF&
    w = (v * x) * 0.0001@
    
    UInt32x16To32 = AsLong(w)
End Function

Private Function UInt32Compare(ByVal x As Long, ByVal y As Long) As Long
    Dim u As Currency
    Dim v As Currency
    
    AsLong(u) = x
    AsLong(v) = y
     
    UInt32Compare = Sgn(u - v)
End Function


Private Sub SinglePlaceMultiply(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, ByRef w() As Integer)
    Dim k As Long
    Dim i As Long
    
    For i = 0 To Length - 1
        k = k + UInt32x16To32(v, u(i))
        w(i) = AsWord(k)
        k = RightShift16(k)
    Next i

    w(Length) = AsWord(k)
End Sub

Private Function MultiInPlaceSubtract(ByRef u() As Integer, ByVal StartIndex As Long, ByRef v() As Integer) As Boolean
    Dim k As Long
    Dim Result As Long
    Dim d As Long
    Dim i As Long
    Dim j As Long
    Dim ubv As Long
    ubv = UBound(v)
    
    For i = StartIndex To UBound(u)
        If j <= ubv Then
            d = v(j) And &HFFFF&
        Else
            d = 0
        End If
        
        Result = Result + ((u(i) And &HFFFF&) - d) + k
        
        If Result < 0 Then
            Result = Result + &H10000
            k = -1
        Else
            k = 0
        End If
        
        u(i) = AsWord(Result)
        Result = RightShift16(Result)
        j = j + 1
    Next i
    
    MultiInPlaceSubtract = k
End Function

Private Sub MultiInPlaceAdd(ByRef u() As Integer, ByVal StartIndex As Long, ByRef v() As Integer)
    Dim Result  As Long
    Dim i       As Long
    Dim j As Long
    Dim d As Long
    Dim ubv As Long
    ubv = UBound(v)
    
    For i = StartIndex To UBound(u)
        If j <= ubv Then
            d = v(j) And &HFFFF&
        Else
            d = 0
        End If
        
        Result = Result + (u(i) And &HFFFF&) + d
        u(i) = AsWord(Result)
        Result = RightShift16(Result)
        j = j + 1
    Next i
End Sub

Public Function SinglePlaceDivide(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, Optional ByRef remainder As Long) As Integer()
    Dim q() As Integer
    ReDim q(0 To Length)
    
    Dim R As Long
    Dim i As Long
    For i = Length - 1 To 0 Step -1
        R = R * &H10000 + (u(i) And &HFFFF&)
        q(i) = AsWord(UInt32d16To32(R, v))
        R = AsWord(UInt32m16To32(R, v))
    Next i
    
    remainder = R
    SinglePlaceDivide = q
End Function

Private Function UInt16x16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim u As Currency
    Dim v As Currency
    Dim w As Currency
    
    u = x And &HFFFF&
    v = y And &HFFFF&
    w = (u * v) * 0.0001@
      
    UInt16x16To32 = AsLong(w)
End Function

Private Function UInt32d16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim d As Currency

    AsLong(d) = x
    d = d * 10000@
    UInt32d16To32 = Int(d / (y And &HFFFF&))
End Function

Private Function UInt32m16To32(ByVal x As Long, ByVal y As Long) As Long
    Dim q As Currency
    Dim d As Currency
    Dim v As Currency
    
    v = y And &HFFFF&
    AsLong(d) = x
    d = d * 10000@
    q = Int(d / v)
    UInt32m16To32 = d - q * v
End Function

Private Function Make32(ByVal x As Integer, ByVal y As Integer) As Long
    Make32 = Helper.ShiftLeft(x And &HFFFF&, 16) Or (y And &HFFFF&)
End Function

Private Function RightShift16(ByVal x As Long) As Long
    RightShift16 = ((x And &HFFFF0000) \ &H10000) And &HFFFF&
End Function
#End If
