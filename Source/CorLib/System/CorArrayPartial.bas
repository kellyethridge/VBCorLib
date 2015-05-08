Attribute VB_Name = "CorArrayPartial"
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
' Module: CorArrayPartial
'

''
' This partial class contains supporting code for the CorArray class.
'
Option Explicit


' Attempt to use a specialized search for a specific data type.
Public Function TrySZBinarySearch(ByVal pSA As Long, ByRef Value As Variant, ByVal StartIndex As Long, ByVal Length As Long, ByRef RetVal As Long) As Boolean
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    RetVal = SZBinarySearch(pSA, VarPtr(CLng(Value)), StartIndex, Length, AddressOf SZCompareLongs)
        Case vbString:                  RetVal = SZBinarySearch(pSA, VarPtr(StrPtr(Value)), StartIndex, Length, AddressOf SZCompareStrings)
        Case vbDouble:                  RetVal = SZBinarySearch(pSA, VarPtr(CDbl(Value)), StartIndex, Length, AddressOf SZCompareDoubles)
        Case vbObject, vbDataObject:    RetVal = SZBinarySearch(pSA, VarPtr(Value), StartIndex, Length, AddressOf SZCompareComparables)
        Case vbInteger:                 RetVal = SZBinarySearch(pSA, VarPtr(CInt(Value)), StartIndex, Length, AddressOf SZCompareIntegers)
        Case vbSingle:                  RetVal = SZBinarySearch(pSA, VarPtr(CSng(Value)), StartIndex, Length, AddressOf SZCompareSingles)
        Case vbCurrency:                RetVal = SZBinarySearch(pSA, VarPtr(CCur(Value)), StartIndex, Length, AddressOf SZCompareCurrencies)
        Case vbDate:                    RetVal = SZBinarySearch(pSA, VarPtr(CDate(Value)), StartIndex, Length, AddressOf SZCompareDates)
        Case vbBoolean:                 RetVal = SZBinarySearch(pSA, VarPtr(CBool(Value)), StartIndex, Length, AddressOf SZCompareBooleans)
        Case vbByte:                    RetVal = SZBinarySearch(pSA, VarPtr(CByte(Value)), StartIndex, Length, AddressOf SZCompareBytes)
        Case Else
            Exit Function
    End Select
    TrySZBinarySearch = True
End Function

' This is an optimized search routine that uses a function pointer
' to call a specific comparison routine.
Private Function SZBinarySearch(ByVal pSA As Long, ByVal pValue As Long, ByVal Index As Long, ByVal Count As Long, ByVal ComparerAddress As Long) As Long
    Dim ElemSize    As Long
    Dim pvData      As Long
    Dim pLowElem    As Long
    Dim pHighElem   As Long
    Dim Comparer    As Func_T_T_Long
    
    ElemSize = SafeArrayGetElemsize(pSA)
    pvData = MemLong(pSA + PVDATA_OFFSET)
    pLowElem = Index - SafeArrayGetLBound(pSA, 1)
    pHighElem = pLowElem + Count - 1
    Set Comparer = NewDelegate(ComparerAddress)
    
    Dim pMiddleElem As Long
    Do While pLowElem <= pHighElem
        pMiddleElem = (pLowElem + pHighElem) \ 2
        Select Case Comparer.Invoke(ByVal pvData + pMiddleElem * ElemSize, ByVal pValue)
            Case 0
                SZBinarySearch = pMiddleElem + SafeArrayGetLBound(pSA, 1)
                Exit Function
            Case Is > 0
                pHighElem = pMiddleElem - 1
            Case Else
                pLowElem = pMiddleElem + 1
        End Select
    Loop
    
    SZBinarySearch = (Not pLowElem) + SafeArrayGetLBound(pSA, 1)
End Function

Private Function SZCompareLongs(ByRef X As Long, ByRef Y As Long) As Long
    If X > Y Then
        SZCompareLongs = 1
    ElseIf X < Y Then
        SZCompareLongs = -1
    End If
End Function

Private Function SZCompareIntegers(ByRef X As Integer, ByRef Y As Integer) As Long
    If X > Y Then
        SZCompareIntegers = 1
    ElseIf X < Y Then
        SZCompareIntegers = -1
    End If
End Function

Private Function SZCompareStrings(ByRef X As String, ByRef Y As String) As Long
    SZCompareStrings = StrComp(X, Y, vbBinaryCompare)
End Function

Private Function SZCompareDoubles(ByRef X As Double, ByRef Y As Double) As Long
    If X > Y Then
        SZCompareDoubles = 1
    ElseIf X < Y Then
        SZCompareDoubles = -1
    End If
End Function

Private Function SZCompareSingles(ByRef X As Single, ByRef Y As Single) As Long
    If X > Y Then
        SZCompareSingles = 1
    ElseIf X < Y Then
        SZCompareSingles = -1
    End If
End Function

Private Function SZCompareBytes(ByRef X As Byte, ByRef Y As Byte) As Long
    If X > Y Then
        SZCompareBytes = 1
    ElseIf X < Y Then
        SZCompareBytes = -1
    End If
End Function

Private Function SZCompareBooleans(ByRef X As Boolean, ByRef Y As Boolean) As Long
    If X > Y Then
        SZCompareBooleans = 1
    ElseIf X < Y Then
        SZCompareBooleans = -1
    End If
End Function

Private Function SZCompareDates(ByRef X As Date, ByRef Y As Date) As Long
    SZCompareDates = DateDiff("s", Y, X)
End Function

Private Function SZCompareCurrencies(ByRef X As Currency, ByRef Y As Currency) As Long
    If X > Y Then
        SZCompareCurrencies = 1
    ElseIf X < Y Then
        SZCompareCurrencies = -1
    End If
End Function

Private Function SZCompareComparables(ByRef X As Object, ByRef Y As Variant) As Long
    Dim XComparable As IComparable
    Set XComparable = X
    SZCompareComparables = XComparable.CompareTo(Y)
End Function




