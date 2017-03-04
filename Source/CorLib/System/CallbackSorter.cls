VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallbackSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: CallbackSorter
'

''
' This class is used to sort arrays using a custom callback compare routine.
'
' @remarks In the sort functions of the CorArray class a function address can be
' passed in as a callback comparer. This class uses that call back to perform
' comparisons of the array elements. This is the second fastest sorting
' provided by VBCorLib. And it allows for complete custom array element comparison.
'
Option Explicit

Private mDelegate       As Delegate
Private mComparer       As Func_T_T_Long
Private mHasSortItems   As Boolean
Private mSortItems      As SortItems
Private mKeyItems       As SortItems

Private mInt8()         As Byte
Private mInt16()        As Integer
Private mInt32()        As Long
Private mInt64()        As Currency
Private mInt128()       As Variant

' used by SortAny
Private mPVData         As Long
Private mElemSize       As Long
Private mPivotBuffer    As Long


Friend Sub Sort(ByVal pSAKeys As Long, ByVal pSAItems As Long, ByVal Left As Long, ByVal Right As Long, ByVal AddressOfCallback As Long)
    mDelegate.pfn = AddressOfCallback
    If pSAItems <> 0 Then
        Call InitSortItems(pSAItems)
    Else
        mHasSortItems = False
    End If
    
    Select Case SafeArrayGetElemsize(pSAKeys)
        Case 1
            SAPtr(mInt8) = pSAKeys
            Call SortInt8(Left, Right)
        Case 2
            SAPtr(mInt16) = pSAKeys
            Call SortInt16(Left, Right)
        Case 4
            SAPtr(mInt32) = pSAKeys
            Call SortInt32(Left, Right)
        Case 8
            SAPtr(mInt64) = pSAKeys
            Call SortInt64(Left, Right)
        Case 16
            SAPtr(mInt128) = pSAKeys
            Call SortInt128(Left, Right)
        Case Else
            With mKeyItems
                Call CopyMemory(.SA, ByVal pSAKeys, vbSizeOfSafeArray1d)
                mElemSize = .SA.cbElements
                mPivotBuffer = CoTaskMemAlloc(mElemSize)
                .Buffer = CoTaskMemAlloc(mElemSize)
                mPVData = .SA.pvData
                Call SortAny(Left, Right)
                Call CoTaskMemFree(.Buffer)
                Call CoTaskMemFree(mPivotBuffer)
            End With
    End Select
    
    If mSortItems.Buffer Then
        Call CoTaskMemFree(mSortItems.Buffer)
        mSortItems.Buffer = 0
    End If
End Sub

Private Sub InitSortItems(ByVal pSA As Long)
    Call CopyMemory(mSortItems.SA, ByVal pSA, vbSizeOfSafeArray1d)
    Select Case mSortItems.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else: mSortItems.Buffer = CoTaskMemAlloc(mSortItems.SA.cbElements)
    End Select
    mHasSortItems = True
End Sub

Private Sub SortInt8(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Byte, t As Byte
    Do While Left < Right
        i = Left: j = Right: x = mInt8((i + j) \ 2)
        Do
            Do While mComparer.Invoke(mInt8(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Invoke(mInt8(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt8(i): mInt8(i) = mInt8(j): mInt8(j) = t: If mHasSortItems Then Call SwapSortItems(mSortItems, i, j)
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call SortInt8(Left, j)
            Left = i
        Else
            If i < Right Then Call SortInt8(i, Right)
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt16(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Integer, t As Integer
    Do While Left < Right
        i = Left: j = Right: x = mInt16((i + j) \ 2)
        Do
            Do While mComparer.Invoke(mInt16(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Invoke(mInt16(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt16(i): mInt16(i) = mInt16(j): mInt16(j) = t: If mHasSortItems Then Call SwapSortItems(mSortItems, i, j)
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call SortInt16(Left, j)
            Left = i
        Else
            If i < Right Then Call SortInt16(i, Right)
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt32(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Long, t As Long
    Do While Left < Right
        i = Left: j = Right: x = mInt32((i + j) \ 2)
        Do
            Do While mComparer.Invoke(mInt32(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Invoke(mInt32(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt32(i): mInt32(i) = mInt32(j): mInt32(j) = t: If mHasSortItems Then Call SwapSortItems(mSortItems, i, j)
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call SortInt32(Left, j)
            Left = i
        Else
            If i < Right Then Call SortInt32(i, Right)
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt64(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Currency, t As Currency
    Do While Left < Right
        i = Left: j = Right: x = mInt64((i + j) \ 2)
        Do
            Do While mComparer.Invoke(mInt64(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Invoke(mInt64(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = mInt64(i): mInt64(i) = mInt64(j): mInt64(j) = t: If mHasSortItems Then Call SwapSortItems(mSortItems, i, j)
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call SortInt64(Left, j)
            Left = i
        Else
            If i < Right Then Call SortInt64(i, Right)
            Right = j
        End If
    Loop
End Sub

Private Sub SortInt128(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, x As Variant
    Do While Left < Right
        i = Left: j = Right: Call CopyMemory(x, mInt128((i + j) \ 2), 16)
        Do
            Do While mComparer.Invoke(mInt128(i), x) < 0: i = i + 1: Loop
            Do While mComparer.Invoke(mInt128(j), x) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Call Helper.Swap16(mInt128(i), mInt128(j)): If mHasSortItems Then Call SwapSortItems(mSortItems, i, j)
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call SortInt128(Left, j)
            Left = i
        Else
            If i < Right Then Call SortInt128(i, Right)
            Right = j
        End If
    Loop
    VariantType(x) = 0
End Sub

Private Sub SortAny(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long: Dim PivotBuffer As Long
    Do While Left < Right
        i = Left: j = Right: Call CopyMemory(ByVal mPivotBuffer, ByVal mPVData + ((i + j) \ 2) * mElemSize, mElemSize)
        Do
            Do While mComparer.Invoke(ByVal mPVData + i * mElemSize, ByVal mPivotBuffer) < 0: i = i + 1: Loop
            Do While mComparer.Invoke(ByVal mPVData + j * mElemSize, ByVal mPivotBuffer) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Call SwapSortItems(mKeyItems, i, j): If mHasSortItems Then Call SwapSortItems(mSortItems, i, j)
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then Call SortAny(Left, j)
            Left = i
        Else
            If i < Right Then Call SortAny(i, Right)
            Right = j
        End If
    Loop
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mComparer = InitDelegate(mDelegate)
End Sub

Private Sub Class_Terminate()
    SAPtr(mInt8) = vbNullPtr
    SAPtr(mInt16) = vbNullPtr
    SAPtr(mInt32) = vbNullPtr
    SAPtr(mInt64) = vbNullPtr
    SAPtr(mInt128) = vbNullPtr
End Sub
