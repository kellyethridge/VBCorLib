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

Private mMaxPath    As Long
Private mCapacity   As Long
Private mCharBuffer As String
Private mChars()    As Integer
Private mLength     As Long
Private mDoNotTryExpandShortFileName As Boolean


Private Sub Class_Initialize()
    mMaxPath = Path.MaxPath
    mCapacity = mMaxPath
    mCharBuffer = String$(mMaxPath, 0)
    mChars = AllocChars(mCharBuffer)
End Sub

Private Sub Class_Terminate()
    FreeChars mChars
End Sub

Friend Property Get Length() As Long
    Length = mLength
End Property

Friend Property Let Length(ByVal Value As Long)
    mLength = Value
End Property

Friend Property Get Capacity() As Long
    Capacity = mCapacity
End Property

Friend Property Get Chars(ByVal Index As Long) As Integer
    If Index < 0 Or Index > mLength Then _
        Error.IndexOutOfRange IndexOutOfRange_ArrayBounds
    
    Chars = mChars(Index)
End Property

Friend Property Let Chars(ByVal Index As Long, ByVal Value As Integer)
    If Index < 0 Or Index > mLength Then _
        Error.IndexOutOfRange IndexOutOfRange_ArrayBounds
        
    Chars(Index) = Value
End Property

Friend Sub Append(ByVal Char As Integer)
    If mLength + 1 > mCapacity Then _
        Error.IOException IOException_PathTooLong
    
    mChars(mLength) = Char
    mLength = mLength + 1
End Sub

Friend Function GetFullPathName() As Long
    Dim FinalBuffer As String
    FinalBuffer = String$(Path.MaxPath + 1, 0)
    
    Dim Result As Long
    Result = GetFullPathNameW(mCharBuffer, mLength, FinalBuffer, ByVal 0)
    If Result >= mMaxPath Then _
        Throw Cor.NewPathTooLongException(Environment.GetResourceString(IOException_PathTooLong))
    If Result = 0 Then _
        IOError Err.LastDllError
    
    CopyMemory mChars(0), ByVal StrPtr(FinalBuffer), Result * 2
    mLength = Result
    GetFullPathName = Result
End Function

Friend Function TryExpandShortFileName() As Boolean
    If mDoNotTryExpandShortFileName Then
        Exit Function
    End If
    
    NullTerminate
    
    Dim ShortFileNameBuffer As String
    ShortFileNameBuffer = String$(Path.MaxPath + 1, 0)
    
    Dim Result As Long
    Result = GetLongPathNameW(mCharBuffer, ShortFileNameBuffer, mMaxPath)
    If Result > mMaxPath Then _
        Throw Cor.NewPathTooLongException(Environment.GetResourceString(IOException_PathTooLong))
    If Result = 0 Then
        Dim DllError As Long
        DllError = Err.LastDllError
        If DllError = ERROR_FILE_NOT_FOUND Or DllError = ERROR_PATH_NOT_FOUND Then
            mDoNotTryExpandShortFileName = True
        End If
        
        Exit Function
    End If
    
    CopyMemory mChars(0), ByVal StrPtr(ShortFileNameBuffer), Result
    mLength = Result
    NullTerminate
    
    TryExpandShortFileName = True
End Function

Friend Sub Fixup(ByVal LenSavedName As Long, ByVal LastSlash As Long)
    Dim SavedName As String
    SavedName = String$(LenSavedName, 0)
    CopyMemory ByVal StrPtr(SavedName), mChars(LastSlash + 1), LenSavedName
    mLength = LastSlash
    NullTerminate
    mDoNotTryExpandShortFileName = False
    
    Dim Result As Boolean
    Result = TryExpandShortFileName
    Append Path.DirectorySeparatorChar
    If mLength + LenSavedName >= mMaxPath Then _
        Throw Cor.NewPathTooLongException(Environment.GetResourceString(IOException_PathTooLong))
    
    CopyMemory mChars(mLength), ByVal StrPtr(SavedName), LenSavedName
    mLength = mLength + LenSavedName
End Sub

Friend Function OrdinalStartsWith(ByRef CompareTo As String, ByVal IgnoreCase As Boolean) As Boolean
    If mLength < Len(CompareTo) Then
        Exit Function
    End If
    
    NullTerminate
    
    Dim Current As String
    Current = ToString
    If IgnoreCase Then
        OrdinalStartsWith = CorString.Equals(Current, CompareTo, OrdinalIgnoreCase)
    Else
        OrdinalStartsWith = CorString.Equals(Current, CompareTo)
    End If
End Function

Public Function ToString() As String
    ToString = Left$(mCharBuffer, mLength)
End Function

Private Sub NullTerminate()
    mChars(mLength) = 0
End Sub


