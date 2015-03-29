VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayValidator"
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
' Module: ArrayValidator
'
Option Explicit

Private mParameter          As Param
Private mIsArray            As Boolean
Private mArrayPtr           As Long
Private mIsMultiDimensional As Boolean
Private mLowerBound         As Long
Private mUpperBound         As Long


Friend Sub ValidateIsArray()
    If Not mIsArray Then
        Throw Cor.NewArgumentException(GetMessage(Argument.ArrayRequired))
    End If
End Sub

Friend Sub ValidateIsOneDimensional()
    If SafeArrayGetDim(mArrayPtr) > 1 Then
        Throw Cor.NewRankException(GetMessage(Argument.MultiDimensionNotSupported))
    End If
End Sub

Friend Sub ValidIndex(ByVal Index As Long, ByVal Parameter As Param)
    If Index < mLowerBound Then
        Throw Cor.NewArgumentOutOfRangeException(GetParameter(Parameter), Index, GetString(ArgumentOutOfRange_ArrayLB))
    End If
End Sub

Friend Sub ValidateOffsetLength(ByVal Index As Long, ByVal Length As Long)
    If Index + Length - 1 >= mUpperBound Then
        Throw Cor.NewArgumentException(GetString(Argument_InvalidOffLen))
    End If
End Sub


Friend Sub Init(ByRef Source As Variant)
    If IsArray(Source) Then
        mIsArray = True
        mArrayPtr = GetArrayPointer(Source)
        
        If mArrayPtr <> vbNullPtr Then
            mIsMultiDimensional = SafeArrayGetDim(mArrayPtr) > 1
            mLowerBound = LBound(Source)
            mUpperBound = UBound(Source)
        End If
    End If
End Sub
