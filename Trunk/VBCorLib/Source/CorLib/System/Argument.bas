Attribute VB_Name = "Argument"
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: ArgumentModule
'

''
' This modules contains functions used to help with method arguments.
'
Option Explicit

' A ListRange is returned from methods that need to return both an Index and Count value.
Public Type ListRange
    Index As Long
    Count As Long
End Type


''
' Returns a pair of optional values, requiring both of them to be missing or present.
'
' @param OptionalValue1 First value of the pair.
' @param DefaultValue1 Default value if the first value is missing.
' @param ReturnValue1 The return parameter of the first value.
' @param OptionalValue2 Second value of the pair.
' @param DefaultValue2 Default value if the second value is missing.
' @param ReturnValue2 The return parameter of the second value.
' @return If the function is successful, then NO_ERROR is returned,
' otherwise, an exception error number is returned.
' @remarks Checks that both optional arguments are either both supplied or both are missing. Cannot supply only one argument.
'
Public Function GetOptionalLongPair(ByRef OptionalValue1 As Variant, ByVal DefaultValue1 As Long, ByRef ReturnValue1 As Long, _
                                    ByRef OptionalValue2 As Variant, ByVal DefaultValue2 As Long, ByRef ReturnValue2 As Long) As Long
    Dim FirstIsMissing As Boolean
    
    FirstIsMissing = IsMissing(OptionalValue1)
    
    If FirstIsMissing = IsMissing(OptionalValue2) Then
        If FirstIsMissing Then
            ReturnValue1 = DefaultValue1
            ReturnValue2 = DefaultValue2
        Else
            ReturnValue1 = CLngOrDefault(OptionalValue1, DefaultValue1)
            ReturnValue2 = CLngOrDefault(OptionalValue2, DefaultValue2)
        End If
    Else
        GetOptionalLongPair = Argument_ParamRequired
    End If
End Function

Public Function MakeArrayRange(ByRef Arr As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As ListRange
    If IsMissing(Index) Then
        MakeArrayRange.Index = LBound(Arr)
    Else
        MakeArrayRange.Index = Index
    End If
    
    If IsMissing(Count) Then
        MakeArrayRange.Count = UBound(Arr) - MakeArrayRange.Index + 1
    Else
        MakeArrayRange.Count = Count
    End If
End Function

Public Function MakeDefaultRange(ByRef Index As Variant, ByRef Count As Variant, ByVal DefaultIndex As Long, ByVal DefaultCount As Long, Optional ByVal IndexName As ParameterName = NameOfIndex, Optional ByVal CountName As ParameterName = NameOfCount) As ListRange
    Dim IndexIsMissing As Boolean
    
    IndexIsMissing = IsMissing(Index)
    
    If IndexIsMissing <> IsMissing(Count) Then _
        Error.Argument Argument_ParamRequired, Environment.GetParameterName(IIf(IndexIsMissing, IndexName, CountName))
    
    If IndexIsMissing Then
        MakeDefaultRange.Index = DefaultIndex
        MakeDefaultRange.Count = DefaultCount
    Else
        MakeDefaultRange.Index = Index
        MakeDefaultRange.Count = Count
    End If
End Function

Public Function MakeDefaultStepRange(ByRef Index As Variant, ByRef Count As Variant, ByVal DefaultIndex As Long, ByVal DefaultCount As Long, Optional ByVal IndexName As ParameterName = NameOfIndex, Optional ByVal CountName As ParameterName = NameOfCount) As ListRange
    If IsMissing(Index) Then
        If Not IsMissing(Count) Then _
            Error.Argument Argument_ParamRequired, Environment.GetParameterName(IndexName)
            
        MakeDefaultStepRange.Index = DefaultIndex
        MakeDefaultStepRange.Count = DefaultCount
    Else
        MakeDefaultStepRange.Index = Index
        MakeDefaultStepRange.Count = CLngOrDefault(Count, DefaultCount - MakeDefaultStepRange.Index)
    End If
End Function

''
' Assigns given values or default values, returning any error codes.
'
' @param pSafeArray A pointer to a SafeArray structure.
' @param OptionalIndex The index value supplied by the caller.
' @param ReturnIndex Returns the index of the starting range of the array.
' @param OptionalCount The count value supplied by the caller.
' @param ReturnCount Returns the number of elements to include in the range.
' @return If the function is successful, then NO_ERROR is returned,
' otherwise, an exception error number is returned.
' @remarks <p>Range checking is performed to ensure a Index and Count value pair do not extend outside of the array.</p>
'
Public Function GetOptionalArrayRange(ByVal pSafeArray As Long, _
                                      ByRef OptionalIndex As Variant, ByRef ReturnIndex As Long, _
                                      ByRef OptionalCount As Variant, ByRef ReturnCount As Long) As Long
    Dim LowerBound As Long
    Dim UpperBound As Long
    
    ' This function is optimized by not refactoring
    ' common sections with other helper rountine in
    ' order to cut down on total function calls.
    
    If pSafeArray = vbNullPtr Then
        GetOptionalArrayRange = ArgumentNull_Array
        Exit Function
    End If
    
    If SafeArrayGetDim(pSafeArray) <> 1 Then
        GetOptionalArrayRange = Rank_MultiDimNotSupported
        Exit Function
    End If
    
    LowerBound = SafeArrayGetLBound(pSafeArray, 1)
    UpperBound = SafeArrayGetUBound(pSafeArray, 1)
    
    Dim Result As Long
    Result = GetOptionalLongPair(OptionalIndex, LowerBound, ReturnIndex, OptionalCount, UpperBound - LowerBound + 1, ReturnCount)
    If Result <> NO_ERROR Then
        GetOptionalArrayRange = Result
        Exit Function
    End If
    
    If ReturnIndex < LowerBound Then
        GetOptionalArrayRange = ArgumentOutOfRange_LBound
        Exit Function
    End If
    
    If ReturnCount < 0 Then
        GetOptionalArrayRange = ArgumentOutOfRange_NeedNonNegNum
        Exit Function
    End If
    
    If ReturnIndex + ReturnCount - 1 > UpperBound Then
        GetOptionalArrayRange = Argument_InvalidOffLen
    End If
End Function


''
' Verifies the index and count are within the bounds and size of a one-dimensional array.
'
' @param pSA A pointer to a SafeArray structure.
' @param Index The index into the array.
' @param Count The number of elements to include.
' @return If this function succeeds, then NO_ERROR is returned, otherwise
' and error exception code is returned.
'
Public Function VerifyArrayRange(ByVal pSafeArray As Long, ByVal Index As Long, ByVal Count As Long) As Long
    ' This function is optimized by not refactoring
    ' common sections with other helper rountine in
    ' order to cut down on total function calls.

    ' Check if the array is a null array.
    If pSafeArray = vbNullPtr Then
        VerifyArrayRange = ArgumentNull_Array
        Exit Function
    End If
    
    ' Ensure we only have a 1-Dimension array.
    If SafeArrayGetDim(pSafeArray) <> 1 Then
        VerifyArrayRange = Rank_MultiDimNotSupported
        Exit Function
    End If
    
    ' Can't have an index before the beginning of the array.
    If Index < SafeArrayGetLBound(pSafeArray, 1) Then
        VerifyArrayRange = ArgumentOutOfRange_LBound
        Exit Function
    End If
    
    ' Can't have a negative count.
    If Count < 0 Then
        VerifyArrayRange = ArgumentOutOfRange_NeedNonNegNum
        Exit Function
    End If
    
    ' Can't have the range extend past the end of the array.
    If Index + Count - 1 > SafeArrayGetUBound(pSafeArray, 1) Then
        VerifyArrayRange = Argument_InvalidOffLen
    End If
End Function

''
' Throws specific exceptions based on an error code.
'
' @param ErrorCode The code that determines which exception to throw.
' @param ArrayName The name of the array in which the exception occurred.
' @param Index The index into the array at the time of the error.
' @param IndexName The name of the Index parameter to be included in the exception.
' @param Count The number of elements that were included in the verification of the range in the array.
' @param CountName The name of the Count parameter to be included in the exception.
' @param IsIndexMissing Used to help determine which parameter was missing in the original function call.
' @remarks This throws exceptions that are general cases about an Index and Count
' being valid within a given array. Not all exception types are represented here.
'
Public Sub ThrowArrayRangeException(ByVal ErrorCode As Long, ByRef ArrayName As String, ByVal Index As Long, ByRef IndexName As String, ByVal Count As Long, ByRef CountName As String, Optional ByVal IsIndexMissing As Boolean)
    Dim Message As String
    Message = Environment.GetResourceString(ErrorCode)
    Select Case ErrorCode
        Case ArgumentNull_Array:                Throw Cor.NewArgumentNullException(ArrayName, Message)
        Case Rank_MultiDimNotSupported:         Throw Cor.NewRankException(Message)
        Case Argument_ParamRequired:            Throw Cor.NewArgumentException(Message, IIf(IsIndexMissing, IndexName, CountName))
        Case ArgumentOutOfRange_LBound:         Throw Cor.NewArgumentOutOfRangeException(IndexName, Message, Index)
        Case ArgumentOutOfRange_UBound:         Throw Cor.NewArgumentOutOfRangeException(IndexName, Message, Index)
        Case ArgumentOutOfRange_NeedNonNegNum:  Throw Cor.NewArgumentOutOfRangeException(CountName, Message, Count)
        Case Argument_InvalidOffLen:            Throw Cor.NewArgumentException(Message, CountName)
        Case Else:                              Throw Cor.NewArgumentException(Message)
    End Select
End Sub
