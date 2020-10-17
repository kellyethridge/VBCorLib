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

Public Type ListRange
    Index As Long
    Count As Long
End Type

Private mShiftedArgumentsSA As SafeArray1d

Public Function ShiftArguments(ByRef Args() As Variant) As Variant()
    Dim Length As Long
    
    Length = Len1D(Args)
    
    With mShiftedArgumentsSA
        .cbElements = vbSizeOfVariant
        .cDims = 1
        .cLocks = 1
        
        If Length > 1 Then
            .PVData = VarPtr(Args(LBound(Args) + 1))
            .cElements = Length - 1
        Else
            .cbElements = 0
        End If
    End With
    
    SAPtr(ShiftArguments) = VarPtr(mShiftedArgumentsSA)
End Function

Public Sub FreeArguments(ByRef Arguments() As Variant)
    mShiftedArgumentsSA.PVData = vbNullPtr
    SAPtr(Arguments) = vbNullPtr
End Sub

Public Function MakeArrayRange(ByRef Arr As Variant, ByRef Index As Variant, ByRef Count As Variant) As ListRange
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

Public Function MakeDefaultRange(ByRef Index As Variant, ByVal DefaultIndex As Long, ByRef Count As Variant, ByVal DefaultCount As Long, Optional ByVal IndexName As ParameterName = NameOfIndex, Optional ByVal CountName As ParameterName = NameOfCount) As ListRange
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

Public Function MakeDefaultStepRange(ByRef Index As Variant, ByVal DefaultIndex As Long, ByRef Count As Variant, ByVal DefaultCount As Long, Optional ByVal IndexName As ParameterName = NameOfIndex) As ListRange
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

