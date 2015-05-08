Attribute VB_Name = "ArrayHelper"
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modArrayHelpers
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

Public Function NewListRange(ByVal Index As Long, ByVal Count As Long) As ListRange
    NewListRange.Index = Index
    NewListRange.Count = Count
End Function

Public Function ArrayLength(ByRef Arr As Variant) As Long
    ArrayLength = UBound(Arr) - LBound(Arr) + 1
End Function

Public Function ArrayPointer(ByRef Arg As Variant) As Long
    Dim ArgType As Integer
    ArgType = VariantType(Arg)
    
    If (ArgType And vbArray) = 0 Then _
        Error.Argument Argument_ArrayRequired
    
    ArrayPointer = MemLong(vbaVarRefAry(Arg))
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any cArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case ArgType And &HFF
        Case vbObject, vbUserDefinedType
            If UBound(Arg) < LBound(Arg) Then
                ArrayPointer = vbNullPtr
            End If
    End Select
End Function

Public Function IsNullArray(ByRef Arr As Variant) As Boolean
    IsNullArray = ArrayPointer(Arr) = vbNullPtr
End Function

Public Function IsMultiDimArray(ByRef Arr As Variant) As Boolean
    IsMultiDimArray = SafeArrayGetDim(ArrayPointer(Arr)) > 1
End Function

Public Function IsNullArrayPtr(ByVal ArrayPtr As Long) As Boolean
    IsNullArrayPtr = ArrayPtr = vbNullPtr
End Function

Public Function IsMultiDimArrayPtr(ByVal ArrayPtr As Long) As Boolean
    IsMultiDimArrayPtr = SafeArrayGetDim(ArrayPtr) > 1
End Function

Public Function ValidArrayPointer(ByRef Arr As Variant, Optional ByRef Parameter As String = "Arr", Optional ByVal Message As ArgumentNullString = ArgumentNull_Array) As Long
    Dim ArrayPtr As Long
    
    ArrayPtr = ArrayPointer(Arr)
    If ArrayPtr = vbNullPtr Then _
        Error.ArgumentNull Parameter, Message
    
    ValidArrayPointer = ArrayPtr
End Function

''
' Retrieves the pointer to an array's SafeArray structure.
'
' @param arr The array to retrieve the pointer to.
' @return A pointer to a SafeArray structure or 0 if the array is null.
'
Public Function GetArrayPointer(ByRef Arr As Variant, Optional ByVal ThrowOnNull As Boolean = False) As Long
    If Not IsArray(Arr) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ArrayRequired), "Arr")
    
    GetArrayPointer = MemLong(vbaVarRefAry(Arr))
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any cArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case VariantType(Arr) And &HFF
        Case vbObject, vbUserDefinedType
            If UBound(Arr) < LBound(Arr) Then
                GetArrayPointer = vbNullPtr
            End If
    End Select
    
    If ThrowOnNull Then
        If GetArrayPointer = vbNullPtr Then
            Throw Cor.NewArgumentNullException("Arr", Environment.GetResourceString(ArgumentNull_Array))
        End If
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Optimized sort routines. There could have been one
'   all-purpose sort routine, but it would be too slow.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetSortKeys(ByVal pSA As Long)
    CopyMemory mSortKeys.SA, ByVal pSA, SizeOfSafeArray1d
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
    CopyMemory mSortItems.SA, ByVal pSA, SizeOfSafeArray1d
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

Public Sub QuickSortLong(ByRef Keys() As Long, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortString(ByRef Keys() As String, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortObject(ByRef Keys() As Object, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortInteger(ByRef Keys() As Integer, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortByte(ByRef Keys() As Byte, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortDouble(ByRef Keys() As Double, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortSingle(ByRef Keys() As Single, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortCurrency(ByRef Keys() As Currency, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortBoolean(ByRef Keys() As Boolean, ByVal Left As Long, ByVal Right As Long)
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

Public Sub QuickSortVariant(ByRef Keys() As Variant, ByVal Left As Long, ByVal Right As Long)
    Dim i As Long, j As Long, X As Variant
    Do While Left < Right
        i = Left: j = Right: VariantCopyInd X, Keys((i + j) \ 2)
        Do
            Do While SZCompareVariants(Keys(i), X) < 0: i = i + 1: Loop
            Do While SZCompareVariants(Keys(j), X) > 0: j = j - 1: Loop
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


