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
' This module contains the primary arithmetic algorithms used on BigNumber type.
'
' When compiling the "Release" compiler constant should be set to 1 and
' "Remove Integer Overflow Checks" should be on. This will compile the optimized
' portions of routines.
'
Option Explicit

''
' This contains all the information about a number. The information can be easily
' passed around as a group instead of trying to pass individual parameters.
Public Type BigNumber
    Digits()    As Integer
    Precision   As Long
    Sign        As Long
End Type


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

Public Function Compare(ByRef x As BigNumber, ByRef y As BigNumber) As Long
    Dim Result As Long
    
    Result = x.Sign - y.Sign
    
    If Result = 0 Then
        Result = x.Precision - y.Precision
        
        If Result = 0 Then
            Dim i As Long
            For i = x.Precision - 1 To 0 Step -1
                Result = (x.Digits(i) And &HFFFF&) - (y.Digits(i) And &HFFFF&)
                If Result <> 0 Then
                    Exit For
                End If
            Next i
        Else
            Result = x.Sign * Result
        End If
    End If
    
    Compare = Sgn(Result)
End Function

Public Sub Negate(ByRef n As BigNumber, ByRef Result As BigNumber)
    ReDim Result.Digits(0 To n.Precision)
    CopyMemory Result.Digits(0), n.Digits(0), n.Precision * 2
    Result.Sign = n.Sign
    Result.Precision = n.Precision
    NegateInPlace Result
End Sub

Private Sub NegateInPlace(ByRef n As BigNumber)
    Dim k As Long
    Dim i As Long
 
    Debug.Assert n.Precision <= UBound(n.Digits)
 
    n.Sign = 0 - n.Sign
    k = 1
    
    For i = 0 To n.Precision - 1
        k = k + ((n.Digits(i) Xor &HFFFF) And &HFFFF&)
        
#If Relese Then
        n.Digits(i) = k And &HFFFF&
#Else
        n.Digits(i) = AsWord(k)
#End If
    
        k = (k And &HFFFF0000) \ vbShift16Bits
    Next

    If n.Sign = -1 Then
        If (n.Digits(n.Precision - 1) And &H8000) = 0 Then
            n.Digits(n.Precision) = &HFFFF
            n.Precision = n.Precision + 1
        End If
    ElseIf n.Sign = 1 Then
        Normalize n
    End If
End Sub

''
' This is the basic implementation of a gradeschool style
' addition of two n-place numbers.
'
' Ref: The Art of Computer Programming 4.3.1.A
'
Public Sub Add(ByRef u As BigNumber, ByRef v As BigNumber, ByRef Result As BigNumber)
    Dim uExtDigit   As Long
    Dim vExtDigit   As Long

    If u.Sign = Negative Then
        uExtDigit = &HFFFF&
    End If
    
    If v.Sign = Negative Then
        vExtDigit = &HFFFF&
    End If

    If u.Precision >= v.Precision Then
        ReDim Result.Digits(0 To u.Precision)
    Else
        ReDim Result.Digits(0 To v.Precision)
    End If

    Dim i As Long
    Dim k As Long
    For i = 0 To UBound(Result.Digits)
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
        
        k = uDigit + vDigit + k
        
#If Release Then
        Result.Digits(i) = k And &HFFFF&
#Else
        Result.Digits(i) = AsWord(k)
#End If
        
        k = (k And &HFFFF0000) \ vbShift16Bits
    Next i
    
    Normalize Result
End Sub

''
' This is the basic implementation of a gradeschool style
' subtraction of two n-place numbers.
'
' Ref: The Art of Computer Programming 4.3.1.S
'
Public Sub Subtract(ByRef u As BigNumber, ByRef v As BigNumber, ByRef Result As BigNumber)
    Dim uExtDigit As Long
    Dim vExtDigit As Long

    If u.Sign = Negative Then
        uExtDigit = &HFFFF&
    End If
    
    If v.Sign = Negative Then
        vExtDigit = &HFFFF&
    End If
    
    If u.Precision > v.Precision Then
        ReDim Result.Digits(0 To u.Precision)
    Else
        ReDim Result.Digits(0 To v.Precision)
    End If
    
    Dim i As Long
    Dim k As Long
    For i = 0 To UBound(Result.Digits)
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

        k = uDigit - vDigit + k
        
#If Release Then
        Result.Digits(i) = k And &HFFFF&
#Else
        Result.Digits(i) = AsWord(k)
#End If

        k = (k And &HFFFF0000) \ vbShift16Bits
    Next i
    
    Normalize Result
End Sub

''
' This is a straight forward implementation of Knuth's algorithm.
'
' Ref: The Art of Computer Programming 4.3.1.M
'
Public Sub Multiply(ByRef u As BigNumber, ByRef v As BigNumber, ByRef Result As BigNumber)
    Debug.Assert u.Sign <> 0
    Debug.Assert v.Sign <> 0
    
    If u.Sign = -1 Then
        If v.Sign = -1 Then
            MultiplyNegatives u, v, Result
        Else
            MultiplyByNegative v, u, Result
        End If
    ElseIf v.Sign = -1 Then
        MultiplyByNegative u, v, Result
    Else
        MultiplyPositives u, v, Result
    End If
End Sub

Private Sub MultiplyNegatives(ByRef n1 As BigNumber, ByRef n2 As BigNumber, ByRef Result As BigNumber)
    Dim u As BigNumber
    Dim v As BigNumber
    
    Debug.Assert n1.Sign = -1
    Debug.Assert n2.Sign = -1
    
    Negate n1, u
    Negate n2, v
    MultiplyPositives u, v, Result
End Sub

Private Sub MultiplyByNegative(ByRef u As BigNumber, ByRef Negative As BigNumber, ByRef Result As BigNumber)
    Dim v As BigNumber
    
    Debug.Assert u.Sign = 1
    Debug.Assert Negative.Sign = -1
    
    Negate Negative, v
    MultiplyPositives u, v, Result
    NegateInPlace Result
End Sub

Private Sub MultiplyPositives(ByRef u As BigNumber, ByRef v As BigNumber, ByRef Result As BigNumber)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim d As Long
    
    Debug.Assert u.Sign = 1
    Debug.Assert v.Sign = 1
    
    ReDim Result.Digits(0 To u.Precision + v.Precision)
    Result.Sign = 1
    Result.Precision = UBound(Result.Digits)
    
    For i = 0 To v.Precision - 1
        k = 0
        d = v.Digits(i) And &HFFFF&
        
        For j = 0 To u.Precision - 1
#If Release Then
            k = d * (u.Digits(j) And &HFFFF&) + (Result.Digits(i + j) And &HFFFF&) + k
            Result.Digits(i + j) = k And &HFFFF&
            k = ((k And &HFFFF0000) \ vbShift16Bits) And &HFFFF&
#Else
            k = UInt16x16To32(d, u.Digits(j)) + (Result.Digits(i + j) And &HFFFF&) + k
            Result.Digits(i + j) = AsWord(k)
            k = RightShift16(k)
#End If
        Next
        
#If Release Then
        Result.Digits(i + j) = k And &HFFFF&
#Else
        Result.Digits(i + j) = AsWord(k)
#End If
    Next
    
    Normalize Result
End Sub

Public Sub Divide(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber)
    Dim Remainder As BigNumber
    DivideNumbers Dividend, Divisor, Quotient, Remainder, False
End Sub

Public Sub DivRem(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber)
    DivideNumbers Dividend, Divisor, Quotient, Remainder, True
End Sub

Public Sub Remainder(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Result As BigNumber)
    Dim Quotient As BigNumber
    DivideNumbers Dividend, Divisor, Quotient, Result, True
    
    If Result.Sign <> 0 Then
        If Result.Sign <> Dividend.Sign Then
            NegateInPlace Result
        End If
    End If
End Sub

Private Sub DivideNumbers(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    If Dividend.Sign = -1 Then
        If Divisor.Sign = -1 Then
            DivideNegatives Dividend, Divisor, Quotient, Remainder, IncludeRemainder
        Else
            DivideNegativeByPositive Dividend, Divisor, Quotient, Remainder, IncludeRemainder
        End If
    ElseIf Divisor.Sign = -1 Then
        DividePositiveByNegative Dividend, Divisor, Quotient, Remainder, IncludeRemainder
    Else
        DividePositives Dividend, Divisor, Quotient, Remainder, IncludeRemainder
    End If
End Sub

Private Sub DivideNegativeByPositive(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    Dim u As BigNumber
    Dim v As BigNumber
    
    Debug.Assert Dividend.Sign = -1
    Debug.Assert Divisor.Sign = 1
    
    Negate Dividend, u
    v = Divisor
    DivideToNegative u, v, Quotient, Remainder, IncludeRemainder
End Sub

Private Sub DividePositiveByNegative(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    Dim u As BigNumber
    Dim v As BigNumber
    
    Debug.Assert Dividend.Sign = 1
    Debug.Assert Divisor.Sign = -1
    
    u = Dividend
    Negate Divisor, v
    DivideToNegative u, v, Quotient, Remainder, IncludeRemainder
End Sub

Private Sub DivideToNegative(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    DivideCore Dividend, Divisor, Quotient, Remainder, IncludeRemainder
    NegateInPlace Quotient
    
    If IncludeRemainder Then
        NegateInPlace Remainder
    End If
End Sub

Private Sub DivideNegatives(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    Dim u As BigNumber
    Dim v As BigNumber
    
    Negate Dividend, u
    Negate Divisor, v
    DivideCore u, v, Quotient, Remainder, IncludeRemainder
End Sub

Private Sub DividePositives(ByRef Dividend As BigNumber, ByRef Divisor As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    Dim u As BigNumber
    Dim v As BigNumber
    
    u = Dividend
    v = Divisor
    DivideCore u, v, Quotient, Remainder, IncludeRemainder
End Sub

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
Private Sub DivideCore(ByRef u As BigNumber, ByRef v As BigNumber, ByRef Quotient As BigNumber, ByRef Remainder As BigNumber, ByVal IncludeRemainder As Boolean)
    Quotient.Sign = 1
    
    If v.Precision = 1 Then
        Dim r As Long
        Quotient.Digits = SinglePlaceDivide(u.Digits, u.Precision, v.Digits(0), r)
        Quotient.Precision = UBound(Quotient.Digits) + 1
        Normalize Quotient
        
        If IncludeRemainder Then
            ReDim Remainder.Digits(1)
            
            If r Then
                Remainder.Digits(0) = r
                Remainder.Sign = 1
                Remainder.Precision = 1
            End If
        End If
        Exit Sub
    End If
    
    Dim n As Long
    Dim m As Long
    Dim d As Long
    
    Debug.Assert u.Sign = 1
    Debug.Assert v.Sign = 1
    
    n = v.Precision
    m = u.Precision - n
     
    ' test if the divisor is shorter than the dividend, if so then just
    ' return a 0 quotient and the dividend as the remainder, if needed.
    If m < 0 Then
        If IncludeRemainder Then
            ReDim Remainder.Digits(0 To u.Precision)
            CopyMemory Remainder.Digits(0), u.Digits(0), u.Precision * 2
        End If
        
        Quotient.Digits = Cor.NewIntegers()
        Exit Sub
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
    
    ReDim Quotient.Digits(0 To m + 1)
    Dim vDigit  As Long
    Dim vDigit2 As Long
    
    ' this is the Vn-1 digit used repeatedly in step D3.
    vDigit = v.Digits(n - 1) And &HFFFF&
    
    ' this is the Vn-2 digit used repeatedly in step D3.
    If n - 2 >= 0 Then
        vDigit2 = v.Digits(n - 2) And &HFFFF&
    End If
    
    Dim qTimesu() As Integer
    ReDim qTimesu(0 To n)
    
#If Release Then
    ' this is an optimistic caching to be used incase
    ' a negative value is encountered. the same value
    ' will always be used regardless, so cache it here.
    Dim q2 As Long
    Dim r2 As Long
    
    q2 = &H7FFFFFFF \ vDigit
    r2 = &H7FFFFFFF - (q2 * vDigit) + 1
#End If

    Dim j       As Long
    Dim rHat    As Long
    Dim qHat    As Long
    For j = m To 0 Step -1
        Dim WordU As Long
        
        ' ** D3 Start **
#If Release Then
        ' since we are shifting left, it is possible that we could turn wordu
        ' into a negative value and will need to deal with it differently later on.
        WordU = ((u.Digits(j + n) And &HFFFF&) * vbShift16Bits) Or (u.Digits(j + n - 1) And &HFFFF&)
        
        ' we have to deal with dividing negatives. They need to work like unsigned.
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
#Else
        WordU = Make32(u.Digits(j + n), u.Digits(j + n - 1))
        qHat = UInt32d16To32(WordU, vDigit)
        rHat = UInt32m16To32(WordU, vDigit)
#End If

        Do
            If qHat < &H10000 Then
#If Release Then
                Dim qHatDigits As Long
                Dim rHatDigits As Long

                qHatDigits = (qHat * vDigit2) '(v.Digits(n - 2) And &HFFFF&))
                rHatDigits = (rHat * &H10000) + (u.Digits(j + n - 2) And &HFFFF&)
                
                If (qHatDigits - &H80000000) <= (rHatDigits - &H80000000) Then
                    Exit Do
                End If
#Else
                If UInt32Compare(UInt32x16To32(qHat, v.Digits(n - 2)), Helper.ShiftLeft(rHat, 16) + (u.Digits(j + n - 2) And &HFFFF&)) <= 0 Then
                    Exit Do
                End If
#End If
            End If
            
            qHat = qHat - 1
            rHat = rHat + (vDigit And &HFFFF&)
        Loop While rHat < &H10000
        ' ** D3 End **
        
        ' ** D4 Start **
        SinglePlaceMultiply v.Digits, n, qHat, qTimesu
        
        Dim Borrowed As Boolean
        Borrowed = MultiInPlaceSubtract(u.Digits, j, qTimesu)
        ' ** D4 End **
        
        ' ** D5 Start **
        If Borrowed Then
            ' ** D6 Start **
            qHat = qHat - 1
            MultiInPlaceAdd u.Digits, j, v.Digits
            ' ** D6 End **
        End If
        ' ** D5 End **
        
        Quotient.Digits(j) = AsWord(qHat)
    Next
    ' ** D2 End **
    
    Normalize Quotient
    
    ' ** D8 Start **
    If IncludeRemainder Then
        If d > 1 Then
            Remainder.Digits = SinglePlaceDivide(u.Digits, n, d)
        Else
            Remainder.Digits = u.Digits
        End If
        
        Normalize Remainder
    End If
    ' ** D8 End **
End Sub

Public Sub BitwiseAnd(ByRef Left As BigNumber, ByRef Right As BigNumber, ByRef Result As BigNumber)
    If Left.Precision >= Right.Precision Then
        BitwiseAndCore Right, Left, Result
    Else
        BitwiseAndCore Left, Right, Result
    End If
End Sub

Private Sub BitwiseAndCore(ByRef ShortNumber As BigNumber, ByRef LongNumber As BigNumber, ByRef Result As BigNumber)
    Dim ExtDigit As Integer
        
    If ShortNumber.Sign = -1 Then
        ExtDigit = &HFFFF
        Result.Precision = LongNumber.Precision
    Else
        Result.Precision = ShortNumber.Precision
    End If
    
    ReDim Result.Digits(0 To Result.Precision)
    
    Dim i As Long
    For i = 0 To ShortNumber.Precision - 1
        Result.Digits(i) = LongNumber.Digits(i) And ShortNumber.Digits(i)
    Next i
            
    For i = ShortNumber.Precision To Result.Precision - 1
        Result.Digits(i) = LongNumber.Digits(i) And ExtDigit
    Next i
    
    If LongNumber.Sign = -1 Then
        Result.Digits(Result.Precision) = &HFFFF And ExtDigit
        Result.Precision = Result.Precision + 1
    End If
    
    Normalize Result
End Sub

Public Sub BitwiseOr(ByRef Left As BigNumber, ByRef Right As BigNumber, ByRef Result As BigNumber)
    If Left.Precision >= Right.Precision Then
        BitwiseOrCore Right, Left, Result
    Else
        BitwiseOrCore Left, Right, Result
    End If
End Sub

Private Sub BitwiseOrCore(ByRef ShortNumber As BigNumber, ByRef LongNumber As BigNumber, ByRef Result As BigNumber)
    Dim ExtDigit As Integer
    
    If ShortNumber.Sign = -1 Then
        ExtDigit = &HFFFF
    End If
    
    ReDim Result.Digits(0 To LongNumber.Precision - 1)
    
    Dim i As Long
    For i = 0 To ShortNumber.Precision - 1
        Result.Digits(i) = LongNumber.Digits(i) Or ShortNumber.Digits(i)
    Next i
    
    For i = ShortNumber.Precision To LongNumber.Precision - 1
        Result.Digits(i) = LongNumber.Digits(i) Or ExtDigit
    Next i
    
    Normalize Result
End Sub

Public Sub BitwiseXor(ByRef Left As BigNumber, ByRef Right As BigNumber, ByRef Result As BigNumber)
    If Left.Precision >= Right.Precision Then
        BitwiseXorCore Right, Left, Result
    Else
        BitwiseXorCore Left, Right, Result
    End If
End Sub

Private Sub BitwiseXorCore(ByRef ShortNumber As BigNumber, ByRef LongNumber As BigNumber, ByRef Result As BigNumber)
    Dim ExtDigit As Integer
    
    If ShortNumber.Sign = -1 Then
        ExtDigit = &HFFFF
    End If
    
    ReDim Result.Digits(0 To LongNumber.Precision)
    
    Dim i As Long
    For i = 0 To ShortNumber.Precision - 1
        Result.Digits(i) = LongNumber.Digits(i) Xor ShortNumber.Digits(i)
    Next i
    
    For i = ShortNumber.Precision To LongNumber.Precision - 1
        Result.Digits(i) = LongNumber.Digits(i) Xor ExtDigit
    Next i
    
    If LongNumber.Sign = -1 Then
        Result.Digits(LongNumber.Precision) = &HFFFF Xor ExtDigit
    End If
    
    Normalize Result
End Sub

Public Sub BitwiseNot(ByRef Value As BigNumber, ByRef Result As BigNumber)
    If Value.Sign = 1 Then
        ReDim Result.Digits(0 To Value.Precision)
        Result.Digits(Value.Precision) = &HFFFF
        Result.Precision = Value.Precision + 1
        Result.Sign = -1
    Else
        ReDim Result.Digits(0 To Value.Precision - 1)
        Result.Precision = Value.Precision
        Result.Sign = 1
    End If
    
    Dim i As Long
    For i = 0 To Value.Precision - 1
        Result.Digits(i) = Not Value.Digits(i)
    Next i
    
    Normalize Result
End Sub

Public Sub Rnd(ByVal Size As Long, ByVal IsNegative As Boolean, ByRef Result As BigNumber)
    Dim WordCount As Long
    
    WordCount = Size \ 2
    ReDim Result.Digits(0 To WordCount)
    Result.Precision = WordCount
    Result.Sign = 1
    
    Dim i As Long
    For i = 0 To WordCount - 1
        Result.Digits(i) = Int(VBA.Rnd * 65536) - 32768
    Next i
    
    If Size And 1 Then
        Result.Digits(WordCount) = VBA.Rnd * 256
    End If
    
    If IsNegative Then
        NegateInPlace Result
    End If
End Sub

Public Sub Normalize(ByRef Number As BigNumber)
    Dim Max As Long
    Dim i   As Long
    
    Max = UBound(Number.Digits)

    Select Case Number.Digits(Max)
        Case 0   ' we have a leading zero digit

            ' now search for the first nonzero digit from the left.
            For i = Max - 1 To 0 Step -1
                If Number.Digits(i) <> 0 Then
                    ' we found a nonzero digit, so set the number
                    Number.Precision = i + 1   ' set the number of digits
                    
                    If Number.Sign = Negative Then
                        If (Number.Digits(i) And &H8000) = 0 Then
                            Number.Digits(Number.Precision) = &HFFFF
                            Number.Precision = Number.Precision + 1
                        End If
                    End If
                    Number.Sign = Positive     ' we know it's positive because of the leading zero
                    Number.Precision = i + 1   ' set the number of digits
                    Exit Sub
                End If
            Next i

            Number.Sign = Zero
            Number.Precision = 0

        Case &HFFFF ' we have a leading negative

            Number.Sign = Negative ' we know this for sure

            For i = Max To 0 Step -1
                If Number.Digits(i) <> &HFFFF Then
                    If Number.Digits(i) And &H8000 Then
                        Number.Precision = i + 1
                    Else
                        Number.Precision = i + 2
                    End If
                    Exit Sub
                End If
            Next i

            ' the array was full of &HFFFF, we only need to represent one.
            Number.Precision = 1

        Case Else
            If Number.Digits(Max) And &H8000 Then
                Number.Sign = Negative
            Else
                Number.Sign = Positive
            End If

            Number.Precision = Max + 1
    End Select
End Sub

''
' Performs a single in-place division by 10, returning the remainder.
'
' The buffer is modified by this routine.
'
Public Function SingleInPlaceDivideBy10(ByRef n As BigNumber) As Long
    Dim r As Long
    Dim i As Long
    Dim f As Boolean
    Dim d As Long

    For i = n.Precision - 1 To 0 Step -1
        r = (r * &H10000) + (n.Digits(i) And &HFFFF&)
        d = r \ 10

#If Release Then
        n.Digits(i) = d And &HFFFF&
#Else
        n.Digits(i) = AsWord(d)
#End If

        r = r - (d * 10)

        If Not f Then
            If n.Digits(i) = 0 Then
                n.Precision = n.Precision - 1
            Else
                f = True
            End If
        End If
    Next

    If n.Precision = 0 Then
        n.Sign = 0
    End If

    SingleInPlaceDivideBy10 = r
End Function

#If Release Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' These Release methods must be compiled with Interger Overflow
' checks off. The methods must also pass all tests once compiled.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
Public Function SinglePlaceDivide(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, Optional ByRef Remainder As Long) As Integer()
    Dim r   As Long
    Dim q() As Integer
    Dim q2  As Long
    Dim r2  As Long
        
    ReDim q(0 To Length)
    v = v And &HFFFF&
    q2 = &H7FFFFFFF \ v
    r2 = &H7FFFFFFF - (q2 * v) + 1
    
    Dim i As Long
    For i = Length - 1 To 0 Step -1
        r = r * &H10000 + (u(i) And &HFFFF&)
        
        If r And &H80000000 Then
            Dim q1 As Long
            q1 = (r And &H7FFFFFFF) \ v
            r = (r And &H7FFFFFFF) - (q1 * v) + r2

            If r >= v Then
                q1 = q1 + 1
                r = r - v
            End If

            q(i) = q1 + q2
        Else
            q(i) = r \ v
            r = r - (q(i) And &HFFFF&) * v
        End If
    Next
    
    Remainder = r
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

Public Function SinglePlaceDivide(ByRef u() As Integer, ByVal Length As Long, ByVal v As Long, Optional ByRef Remainder As Long) As Integer()
    Dim q() As Integer
    ReDim q(0 To Length)
    
    Dim r As Long
    Dim i As Long
    For i = Length - 1 To 0 Step -1
        r = r * vbShift16Bits + (u(i) And &HFFFF&)
        q(i) = AsWord(UInt32d16To32(r, v))
        r = AsWord(UInt32m16To32(r, v))
    Next i
    
    Remainder = r
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
