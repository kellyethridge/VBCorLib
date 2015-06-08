Attribute VB_Name = "CorArraySort"
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
' Module: CorArraySort
'
Option Explicit

Private Declare Function vbaVarRefAry Lib "msvbvm60.dll" Alias "__vbaRefVarAry" (ByRef ArrayToDeref As Variant) As Long

Public Type SortItems
    SA      As SafeArray1d
    Buffer  As Long
End Type

Private mSortItems      As SortItems
Private mHasSortItems   As Boolean
Private mSortKeys       As SortItems
Public SortComparer     As IComparer



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Optimized sort routines. There could have been one
'   all-purpose sort routine, but it would be too slow.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TrySZSort(ByVal pSA As Long, ByVal Left As Long, ByVal Right As Long) As Boolean
    Dim pfn As Long
    Select Case SafeArrayGetVartype(pSA) And &HFF
        Case vbLong:                    pfn = FuncAddr(AddressOf QuickSortLong)
        Case vbString:                  pfn = FuncAddr(AddressOf QuickSortString)
        Case vbDouble, vbDate:          pfn = FuncAddr(AddressOf QuickSortDouble)
        Case vbObject, vbDataObject:    pfn = FuncAddr(AddressOf QuickSortObject)
        Case vbVariant:                 pfn = FuncAddr(AddressOf QuickSortVariant)
        Case vbInteger:                 pfn = FuncAddr(AddressOf QuickSortInteger)
        Case vbSingle:                  pfn = FuncAddr(AddressOf QuickSortSingle)
        Case vbByte:                    pfn = FuncAddr(AddressOf QuickSortByte)
        Case vbCurrency:                pfn = FuncAddr(AddressOf QuickSortCurrency)
        Case vbBoolean:                 pfn = FuncAddr(AddressOf QuickSortBoolean)
        Case Else: Exit Function
    End Select
    
    Dim Sorter As Action_T_T_T
    Set Sorter = NewDelegate(pfn)
    Sorter.Invoke pSA, ByVal Left, ByVal Right

    TrySZSort = True
End Function

Public Sub SetSortKeys(ByVal pSA As Long)
    CopyMemory mSortKeys.SA, ByVal pSA, corSizeOfSafeArray1d
    Select Case mSortKeys.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else: mSortKeys.Buffer = CoTaskMemAlloc(mSortKeys.SA.cbElements)
    End Select
End Sub

Public Sub ClearSortKeys()
    If mSortKeys.Buffer Then
        CoTaskMemFree mSortKeys.Buffer
    End If
    mSortKeys.Buffer = 0
End Sub

Public Sub SetSortItems(ByVal pSA As Long)
    CopyMemory mSortItems.SA, ByVal pSA, corSizeOfSafeArray1d
    Select Case mSortItems.SA.cbElements
        Case 1, 2, 4, 8, 16
        Case Else
            mSortItems.Buffer = CoTaskMemAlloc(mSortItems.SA.cbElements)
    End Select
    mHasSortItems = True
End Sub

Public Sub ClearSortItems()
    If mHasSortItems Then
        CoTaskMemFree mSortItems.Buffer
        mSortItems.Buffer = 0
        mHasSortItems = False
    End If
End Sub

Public Sub SwapSortItems(ByRef Items As SortItems, ByVal i As Long, ByVal j As Long)
    With Items.SA
        Select Case .cbElements
            Case 1:     Helper.Swap1 ByVal .pvData + i, ByVal .pvData + j
            Case 2:     Helper.Swap2 ByVal .pvData + i * 2, ByVal .pvData + j * 2
            Case 4:     Helper.Swap4 ByVal .pvData + i * 4, ByVal .pvData + j * 4
            Case 8:     Helper.Swap8 ByVal .pvData + i * 8, ByVal .pvData + j * 8
            Case 16:    Helper.Swap16 ByVal .pvData + i * 16, ByVal .pvData + j * 16
            Case Else
                ' primarily for UDTs
                CopyMemory ByVal Items.Buffer, ByVal .pvData + i * .cbElements, .cbElements
                CopyMemory ByVal .pvData + i * .cbElements, ByVal .pvData + j * .cbElements, .cbElements
                CopyMemory ByVal .pvData + j * .cbElements, ByVal Items.Buffer, .cbElements
        End Select
    End With
End Sub

Private Sub QuickSortLong(ByRef Keys() As Long, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Long, t As Long
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortLong Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortLong Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortString(ByRef Keys() As String, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As String
    Do While Left < Right
        i = Left: j = Right: X = StringRef(Keys((i + j) \ 2))
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap4 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortString Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortString Keys, i, Right
            Right = j
        End If
        StringPtr(X) = 0
    Loop
End Sub

Private Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Variant, Key As IComparable
    Do While Left < Right
        i = Left: j = Right: Set X = Keys((i + j) \ 2)
        Do
            Set Key = Keys(i): Do While Key.CompareTo(X) < 0: i = i + 1: Set Key = Keys(i): Loop
            Set Key = Keys(j): Do While Key.CompareTo(X) > 0: j = j - 1: Set Key = Keys(j): Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap4 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortObject Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortObject Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortInteger(ByRef Keys() As Integer, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Integer, t As Integer
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortInteger Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortInteger Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortByte(ByRef Keys() As Byte, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Byte, t As Byte
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortByte Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortByte Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortDouble(ByRef Keys() As Double, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Double, t As Double
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortDouble Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortDouble Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortSingle(ByRef Keys() As Single, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Single, t As Single
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortSingle Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortSingle Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortCurrency(ByRef Keys() As Currency, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Currency, t As Currency
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortCurrency Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortCurrency Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortBoolean(ByRef Keys() As Boolean, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Boolean, t As Boolean
    Do While Left < Right
        i = Left: j = Right: X = Keys((i + j) \ 2)
        Do
            Do While Keys(i) < X: i = i + 1: Loop
            Do While Keys(j) > X: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then t = Keys(i): Keys(i) = Keys(j): Keys(j) = t: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortBoolean Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortBoolean Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Private Sub QuickSortVariant(ByRef Keys() As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd X, Keys((i + j) \ 2)
        Do
            Do While Comparer.Default.Compare(Keys(i), X) < 0: i = i + 1: Loop
            Do While Comparer.Default.Compare(Keys(j), X) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then Helper.Swap16 Keys(i), Keys(j): If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortVariant Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortVariant Keys, i, Right
            Right = j
        End If
    Loop
End Sub

Public Sub QuickSortGeneral(ByRef Keys As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd X, Keys((i + j) \ 2)
        Do
            Do While SortComparer.Compare(Keys(i), X) < 0: i = i + 1: Loop
            Do While SortComparer.Compare(Keys(j), X) > 0: j = j - 1: Loop
            If i > j Then Exit Do
            If i < j Then SwapSortItems mSortKeys, i, j: If mHasSortItems Then SwapSortItems mSortItems, i, j
            i = i + 1: j = j - 1
        Loop While i <= j
        If j - Left <= Right - i Then
            If Left < j Then QuickSortGeneral Keys, Left, j
            Left = i
        Else
            If i < Right Then QuickSortGeneral Keys, i, Right
            Right = j
        End If
    Loop
End Sub



