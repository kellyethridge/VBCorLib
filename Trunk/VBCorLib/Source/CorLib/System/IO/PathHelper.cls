VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathHelper"
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
' Module: PathHelper
'
Option Explicit

Private mFinalBuffer()  As Integer
Private mPosition       As Long
Private mPrefixSize     As Long
Private mStopExpanding  As Boolean

Friend Sub Append(ByVal Ch As Integer)
    If mPosition > UBound(mFinalBuffer) Then
        ReDim Preserve mFinalBuffer(0 To mPosition * 2)
    End If

    mFinalBuffer(mPosition) = Ch
    mPosition = mPosition + 1
End Sub

Friend Sub SetVolumeLabel(ByVal VolumeChar As Integer)
    InitExtendedPathLength
    Append VolumeChar
    Append vbVolumeSeparatorChar
End Sub

Friend Property Get Chars(ByVal Index As Long) As Integer
    Debug.Assert Index >= mPrefixSize And Index < (mPosition - mPrefixSize)
    
    Chars = mFinalBuffer(Index + mPrefixSize)
End Property

Friend Function GetFullPathName() As Long
    Dim NewBuffer() As Integer
    Dim NewLength   As Long
    Dim AppendSlash As Boolean
    
    ReDim NewBuffer(0 To mPosition + MAX_PATH)
    
    If Length = 4 Then
        AppendSlash = mFinalBuffer(mPosition - 2) = vbDirectorySeparatorChar And _
                      mFinalBuffer(mPosition - 1) = vbPeriodChar
    End If
    
    NewLength = GetFullPathNameW(mFinalBuffer(0), mPosition + MAX_PATH, NewBuffer(0), ByVal vbNullPtr)
    Helper.Swap4 ByVal ArrPtr(mFinalBuffer), ByVal ArrPtr(NewBuffer)
    mPosition = NewLength
    If AppendSlash Then
        Append vbDirectorySeparatorChar
    End If
    
    GetFullPathName = Length
End Function

Friend Function TryExpandShortFileName() As Boolean
    If mStopExpanding Then
        Exit Function
    End If
    
    Dim NewBuffer() As Integer
    Dim Result      As Long
    
    ReDim NewBuffer(0 To mPosition + MAX_PATH)
    
    Result = GetLongPathNameW(mFinalBuffer(0), NewBuffer(0), mPosition + MAX_PATH)
    
    If Result <> 0 Then
        Helper.Swap4 ByVal ArrPtr(mFinalBuffer), ByVal ArrPtr(NewBuffer)
        mPosition = Result
        TryExpandShortFileName = True
    Else
        If Result = ERROR_FILE_NOT_FOUND Or Result = ERROR_PATH_NOT_FOUND Then
            mStopExpanding = True
        End If
    End If
End Function

Friend Property Get Length() As Long
    Length = mPosition - mPrefixSize
End Property

Friend Function ToString() As String
    ToString = SysAllocStringLen(VarPtr(mFinalBuffer(mPrefixSize)), Length)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ReDim mFinalBuffer(0 To MAX_PATH)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitExtendedPathLength()
    Append vbDirectorySeparatorChar
    Append vbDirectorySeparatorChar
    Append vbQuestionMarkChar
    Append vbDirectorySeparatorChar
    mPrefixSize = 4
End Sub
