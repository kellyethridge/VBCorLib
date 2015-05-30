Attribute VB_Name = "ArgumentHelper"
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: modArgumentHelpers
'

''
' This modules contains functions used to help with optional parameters and
' verifying ranges of values.
' <p>All of the functions return an error code except OptionalLong. This
' function returns a valid integer value or throws an exception if a the
' supplied optional value is not an integer type.</p>
'
Option Explicit

Public Type ListRange
    Index As Long
    Count As Long
End Type


''
' Returns an optional value or a default value if the optional value is missing.
'
Public Function OptionalLong(ByRef Value As Variant, ByVal Default As Long) As Long
    Select Case VarType(Value)
        Case vbMissing
            OptionalLong = Default
        Case vbLong, vbInteger, vbByte
            OptionalLong = Value
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(InvalidCast_FromTo, TypeName(Value), "Long"))
    End Select
End Function

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
            ReturnValue1 = OptionalLong(OptionalValue1, DefaultValue1)
            ReturnValue2 = OptionalLong(OptionalValue2, DefaultValue2)
        End If
    Else
        GetOptionalLongPair = Argument_ParamRequired
    End If
End Function

Public Function OptionalRange(ByRef Index As Variant, ByRef Count As Variant, ByVal DefaultIndex As Long, ByVal DefaultCount As Long, Optional ByVal IndexParameter As ResourceStringKey = Parameter_Index, Optional ByVal CountParameter As ResourceStringKey = Parameter_Count) As ListRange
    Dim IndexIsMissing As Boolean
    
    IndexIsMissing = IsMissing(Index)
    
    If IndexIsMissing <> IsMissing(Count) Then _
        Throw Error.Argument(Argument_ParamRequired, IIf(IndexIsMissing, "Index", "Count"))
    
    If IndexIsMissing Then
        OptionalRange.Index = DefaultIndex
        OptionalRange.Count = DefaultCount
    Else
        OptionalRange.Index = OptionalLong(Index, DefaultIndex)
        OptionalRange.Count = OptionalLong(Count, DefaultCount)
    End If
End Function

Public Function OptionalStepRange(ByRef Index As Variant, ByRef Count As Variant, ByVal DefaultIndex As Long, ByVal DefaultCount As Long, Optional ByVal IndexParameter As ResourceStringKey = Parameter_Index, Optional ByVal CountParameter As ResourceStringKey = Parameter_Count) As ListRange
    If IsMissing(Index) Then
        OptionalStepRange.Index = DefaultIndex
        
        If IsMissing(Count) Then
            OptionalStepRange.Count = DefaultCount
        Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), Environment.GetResourceString(IndexParameter))
        End If
    Else
        OptionalStepRange.Index = CLng(Index)
        OptionalStepRange.Count = OptionalLong(Count, DefaultCount - OptionalStepRange.Index)
    End If
End Function

Public Function OptionalReverseStepRange(ByRef Index As Variant, ByRef Count As Variant, ByVal DefaultIndex As Long, ByVal DefaultCount As Long, Optional ByVal IndexParameter As ResourceStringKey = Parameter_Index, Optional ByVal CountParameter As ResourceStringKey = Parameter_Count) As ListRange
    If IsMissing(Index) Then
        OptionalReverseStepRange.Index = DefaultIndex
        
        If IsMissing(Count) Then
            OptionalReverseStepRange.Count = DefaultCount
        Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), Environment.GetResourceString(IndexParameter))
        End If
    Else
        OptionalReverseStepRange.Index = CLng(Index)
        OptionalReverseStepRange.Count = OptionalLong(Count, OptionalReverseStepRange.Index + 1)
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
        Case ArgumentOutOfRange_LBound:         Throw Cor.NewArgumentOutOfRangeException(IndexName, Index, Message)
        Case ArgumentOutOfRange_UBound:         Throw Cor.NewArgumentOutOfRangeException(IndexName, Index, Message)
        Case ArgumentOutOfRange_NeedNonNegNum:  Throw Cor.NewArgumentOutOfRangeException(CountName, Count, Message)
        Case Argument_InvalidOffLen:            Throw Cor.NewArgumentException(Message, CountName)
        Case Else:                              Throw Cor.NewArgumentException(Message)
    End Select
End Sub

''
' Verifies the Index and Count range remain inside the bounds of a 0-base list.
'
' @param ListSize The size of the list being checked against.
' @param Index The index into the list, starting at 0.
' @param Count The number of elements in the list to include in the verification.
' @return If success, then 0 is returned, otherwise an error code is returned.
'
Public Function VerifyListRange(ByVal RangeSize As Long, ByVal Index As Long, ByVal Count As Long) As Long
    If Index < 0 Then
        VerifyListRange = ArgumentOutOfRange_LBound ' this should be mapped to ArgumentOutOfRange_NeedNonNegNum
        Exit Function
    End If

    If Count < 0 Then
        VerifyListRange = ArgumentOutOfRange_NeedNonNegNum
        Exit Function
    End If

    If Index + Count > RangeSize Then
        VerifyListRange = Argument_InvalidOffLen
    End If
End Function

''
' Throws a specific list exception based on the error code, providing the correct argument in the exception.
'
' @param ErrorCode The code of the exception to be thrown.
' @param ListName The name of the list or type of list that was used in verification.
' @param Index The index into the list.
' @param IndexName The name of the variable in the public function.
' @param Count The number of elements verified against the list.
' @param CountName The name of the count variable in the public function.
' @param IsIndexMissing Flag indicating if Index was a missing argument, otherwise Count was missing.
'
Public Sub ThrowListRangeException(ByVal ErrorCode As Long, ByVal Index As Long, ByRef IndexName As String, ByVal Count As Long, ByRef CountName As String, Optional ByVal IsIndexMissing As Boolean)
    Dim Message As String
    
    ' Two of the error codes are of the same type, but we need to distinguish
    ' between them. So we use ArgumentOutOfRange_LBound for one of the codes.
    ' Now get the message for the original error code instead.
    If ErrorCode = ArgumentOutOfRange_LBound Then
        Message = Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum)
    Else
        Message = Environment.GetResourceString(ErrorCode)
    End If
    
    Select Case ErrorCode
        Case Argument_ParamRequired:                Throw Cor.NewArgumentException(Message, IIf(IsIndexMissing, IndexName, CountName))
        Case ArgumentOutOfRange_LBound:             Throw Cor.NewArgumentOutOfRangeException(IndexName, Index, Message)
        Case ArgumentOutOfRange_NeedNonNegNum:      Throw Cor.NewArgumentOutOfRangeException(CountName, Count, Message)
        Case Argument_InvalidOffLen:           Throw Cor.NewArgumentException("The index plus the count extends past the end of the list or string", "Count")
    End Select
End Sub

Public Sub ThrowMissing(ByRef ParameterToCheck As Variant, ByVal FirstParameter As ResourceStringKey, ByVal SecondParameter As ResourceStringKey)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ParamRequired), IIf(IsMissing(ParameterToCheck), Environment.GetResourceString(FirstParameter), Environment.GetResourceString(SecondParameter)))
End Sub

''
' Assigns values to the missing Index and Count pair, returning any error codes.
'
' @param RangeSize The number of elements in the range to verify against.
' @param OptionalIndex The optional Index variable in the public interface.
' @param ReturnIndex The return value for Index.
' @param OptionalCount The optional Count variable in the public interface.
' @param ReturnCount The return value for Count.
' @return Returns NO_ERROR if successful, otherwise an error code is returned.
' @remarks This has a hardcoded lowerbound of 0 for the start of the list.
'
Public Function GetOptionalListRange(ByVal RangeSize As Long, _
                                     ByRef OptionalIndex As Variant, ByRef ReturnIndex As Long, _
                                     ByRef OptionalCount As Variant, ByRef ReturnCount As Long) As Long
    ' This function is optimized by not refactoring
    ' common sections with other helper rountine in
    ' order to cut down on total function calls.
    
    ' Get our optional values.
    Dim Result As Long
    Result = GetOptionalLongPair(OptionalIndex, 0, ReturnIndex, OptionalCount, RangeSize, ReturnCount)
    If Result <> NO_ERROR Then
        GetOptionalListRange = Result
        Exit Function
    End If
    
    ' Can't have an index before the start of the range.
    If ReturnIndex < 0 Then
        GetOptionalListRange = ArgumentOutOfRange_LBound    ' this should be mapped to ArgumentOutOfRange_NeedNonNegNum
        Exit Function
    End If
    
    ' Can't have a negative count.
    If ReturnCount < 0 Then
        GetOptionalListRange = ArgumentOutOfRange_NeedNonNegNum
        Exit Function
    End If
    
    ' Can't have the range extend past the beginning of the array.
    If ReturnIndex + ReturnCount > RangeSize Then
        GetOptionalListRange = Argument_InvalidOffLen
    End If
End Function


