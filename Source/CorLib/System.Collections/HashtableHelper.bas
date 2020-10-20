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
'@Folder("CorLib.System.Collections")
Option Explicit

Public Enum HashBucketState
    EmptyState
    OccupiedState
    DeletedState
End Enum

Public Type HashBucket
    Key         As Variant
    Value       As Variant
    HashCode    As Long
    State       As HashBucketState
End Type

Private mCapacities(0 To 71)    As Long
Private mInited                 As Boolean


''
' Returns the next prime number equal to or above the requested size.
'
Public Function GetHashtableCapacity(ByVal Value As Long) As Long
    If Not mInited Then
        InitPrimes
    End If
        
    ' we'll do a very fast binary search locally.
    Dim High    As Long
    Dim Low     As Long
    Dim Index   As Long
    
    High = 71
    Do While Low <= High
        Index = (Low + High) \ 2
        Select Case mCapacities(Index)
            Case Value
                GetHashtableCapacity = Value
                Exit Function
            Case Is > Value
                High = Index - 1
            Case Else
                Low = Index + 1
        End Select
    Loop
    
    If Index < 0 Then
        Index = Not Index
    End If
    
    GetHashtableCapacity = mCapacities(Index)
End Function

Private Sub InitPrimes()
   mCapacities(0) = 13
   mCapacities(1) = 17
   mCapacities(2) = 23
   mCapacities(3) = 29
   mCapacities(4) = 41
   mCapacities(5) = 53
   mCapacities(6) = 67
   mCapacities(7) = 89
   mCapacities(8) = 113
   mCapacities(9) = 149
   mCapacities(10) = 191
   mCapacities(11) = 251
   mCapacities(12) = 317
   mCapacities(13) = 409
   mCapacities(14) = 541
   mCapacities(15) = 691
   mCapacities(16) = 907
   mCapacities(17) = 1171
   mCapacities(18) = 1523
   mCapacities(19) = 1973
   mCapacities(20) = 2557
   mCapacities(21) = 3323
   mCapacities(22) = 4327
   mCapacities(23) = 5623
   mCapacities(24) = 7283
   mCapacities(25) = 9461
   mCapacities(26) = 12289
   mCapacities(27) = 15971
   mCapacities(28) = 20743
   mCapacities(29) = 26947
   mCapacities(30) = 35023
   mCapacities(31) = 45481
   mCapacities(32) = 59029
   mCapacities(33) = 76673
   mCapacities(34) = 99607
   mCapacities(35) = 129379
   mCapacities(36) = 168067
   mCapacities(37) = 218287
   mCapacities(38) = 283553
   mCapacities(39) = 368323
   mCapacities(40) = 478427
   mCapacities(41) = 621451
   mCapacities(42) = 807241
   mCapacities(43) = 1048583
   mCapacities(44) = 1362059
   mCapacities(45) = 1769281
   mCapacities(46) = 2298209
   mCapacities(47) = 2985287
   mCapacities(48) = 3877763
   mCapacities(49) = 5037091
   mCapacities(50) = 6542959
   mCapacities(51) = 8499037
   mCapacities(52) = 11039929
   mCapacities(53) = 14340433
   mCapacities(54) = 18627667
   mCapacities(55) = 24196619
   mCapacities(56) = 31430473
   mCapacities(57) = 40826971
   mCapacities(58) = 53032703
   mCapacities(59) = 68887367
   mCapacities(60) = 89482037
   mCapacities(61) = 116233673
   mCapacities(62) = 150983087
   mCapacities(63) = 196121153
   mCapacities(64) = 254753797
   mCapacities(65) = 330915313
   mCapacities(66) = 429846191
   mCapacities(67) = 558353591
   mCapacities(68) = 725279729
   mCapacities(69) = 942110419
   mCapacities(70) = 1223764877
   mCapacities(71) = 2147483647
   mInited = True
End Sub

