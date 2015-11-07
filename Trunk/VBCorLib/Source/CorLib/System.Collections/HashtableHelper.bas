Attribute VB_Name = "HashtableHelper"
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
' Module: HashtableHelper
'
Option Explicit

Public Enum BucketStateEnum
    bsEmpty
    bsOccupied
    bsDeleted
End Enum

Public Type Bucket
    Key         As Variant
    Value       As Variant
    HashCode    As Long
    State       As BucketStateEnum
End Type

Private mPrimes(0 To 71)    As Long
Private mInited             As Boolean


''
' Returns the next prime number equal to or above the requested size.
'
Public Function GetPrime(ByVal Value As Long) As Long
    If Not mInited Then
        InitPrimes
    End If
        
    ' we'll do a very fast binary search locally.
    Dim hi As Long
    Dim Lo As Long
    Dim md As Long
    hi = 71
    Do While Lo <= hi
        md = (Lo + hi) \ 2
        Select Case mPrimes(md)
            Case Value
                GetPrime = Value
                Exit Function
            Case Is > Value
                hi = md - 1
            Case Else
                Lo = md + 1
        End Select
    Loop
    
    If md < 0 Then
        md = Not md
    End If
    
    GetPrime = mPrimes(md)
End Function

Private Sub InitPrimes()
   mPrimes(0) = 13
   mPrimes(1) = 17
   mPrimes(2) = 23
   mPrimes(3) = 29
   mPrimes(4) = 41
   mPrimes(5) = 53
   mPrimes(6) = 67
   mPrimes(7) = 89
   mPrimes(8) = 113
   mPrimes(9) = 149
   mPrimes(10) = 191
   mPrimes(11) = 251
   mPrimes(12) = 317
   mPrimes(13) = 409
   mPrimes(14) = 541
   mPrimes(15) = 691
   mPrimes(16) = 907
   mPrimes(17) = 1171
   mPrimes(18) = 1523
   mPrimes(19) = 1973
   mPrimes(20) = 2557
   mPrimes(21) = 3323
   mPrimes(22) = 4327
   mPrimes(23) = 5623
   mPrimes(24) = 7283
   mPrimes(25) = 9461
   mPrimes(26) = 12289
   mPrimes(27) = 15971
   mPrimes(28) = 20743
   mPrimes(29) = 26947
   mPrimes(30) = 35023
   mPrimes(31) = 45481
   mPrimes(32) = 59029
   mPrimes(33) = 76673
   mPrimes(34) = 99607
   mPrimes(35) = 129379
   mPrimes(36) = 168067
   mPrimes(37) = 218287
   mPrimes(38) = 283553
   mPrimes(39) = 368323
   mPrimes(40) = 478427
   mPrimes(41) = 621451
   mPrimes(42) = 807241
   mPrimes(43) = 1048583
   mPrimes(44) = 1362059
   mPrimes(45) = 1769281
   mPrimes(46) = 2298209
   mPrimes(47) = 2985287
   mPrimes(48) = 3877763
   mPrimes(49) = 5037091
   mPrimes(50) = 6542959
   mPrimes(51) = 8499037
   mPrimes(52) = 11039929
   mPrimes(53) = 14340433
   mPrimes(54) = 18627667
   mPrimes(55) = 24196619
   mPrimes(56) = 31430473
   mPrimes(57) = 40826971
   mPrimes(58) = 53032703
   mPrimes(59) = 68887367
   mPrimes(60) = 89482037
   mPrimes(61) = 116233673
   mPrimes(62) = 150983087
   mPrimes(63) = 196121153
   mPrimes(64) = 254753797
   mPrimes(65) = 330915313
   mPrimes(66) = 429846191
   mPrimes(67) = 558353591
   mPrimes(68) = 725279729
   mPrimes(69) = 942110419
   mPrimes(70) = 1223764877
   mPrimes(71) = 2147483647
   mInited = True
End Sub

