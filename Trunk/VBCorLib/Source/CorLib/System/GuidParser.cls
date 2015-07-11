VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuidParser"
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
' Module: GuidParser
'
Option Explicit

Private Enum ParseOptions
    None = 0
    Dashes = 1
    Braces = 2
    Parens = 4
End Enum

Private mException As Exception

Friend Sub ThrowError()
    Throw mException
End Sub

Friend Function TryParseExact(ByRef s As String, ByVal Format As Integer, ByRef Result As VBGUID) As Boolean
    Dim GuidString  As String
    Dim GuidChars() As Integer
    
    GuidString = CorString.Trim(s)
    GuidChars = AllocChars(GuidString)
        
    Select Case Format
        Case corLowerBChar, corUpperBChar
            If Len(GuidString) = 38 Then
                TryParseExact = TryParseCore(GuidChars, Dashes Or Braces, Result)
            Else
                FormatError
            End If
            
        Case corLowerPChar, corUpperPChar
            If Len(GuidString) = 38 Then
                TryParseExact = TryParseCore(GuidChars, Dashes Or Parens, Result)
            Else
                FormatError
            End If
            
        Case corLowerNChar, corUpperNChar
            If Len(GuidString) = 32 Then
                TryParseExact = TryParseCore(GuidChars, ParseOptions.None, Result)
            Else
                FormatError
            End If
            
        Case corLowerDChar, corUpperDChar
            If Len(GuidString) = 36 Then
                TryParseExact = TryParseCore(GuidChars, Dashes, Result)
            Else
                FormatError
            End If
            
        Case Else
            FormatSpecificationError
    End Select
    
    FreeChars GuidChars
End Function

Private Function TryParseCore(ByRef Chars() As Integer, ByVal Options As ParseOptions, ByRef Result As VBGUID) As Boolean
    Dim Index As Long
    
    If Options And Braces Then
        If Chars(Index) <> corOpenBraceChar Then
            FormatError
            Exit Function
        End If
        Index = Index + 1
    ElseIf Options And Parens Then
        If Chars(Index) <> corOpenParenChar Then
            FormatError
            Exit Function
        End If
        Index = Index + 1
    End If
    
    If Not TryParseHex(Chars(), Index, 8, Result.Data1) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryParseDash(Chars, Index) Then
            Exit Function
        End If
    End If
    
    If Not TryParseHexInt16(Chars, Index, Result.Data2) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryParseDash(Chars, Index) Then
            Exit Function
        End If
    End If
    
    If Not TryParseHexInt16(Chars, Index, Result.Data3) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryParseDash(Chars, Index) Then
            Exit Function
        End If
    End If
    
    If Not TryParseHexByte(Chars, Index, Result.Data4(0)) Then
        Exit Function
    End If
    If Not TryParseHexByte(Chars, Index, Result.Data4(1)) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryParseDash(Chars, Index) Then
            Exit Function
        End If
    End If
    
    If Not TryParseHexByte(Chars, Index, Result.Data4(2)) Then
        Exit Function
    End If
    If Not TryParseHexByte(Chars, Index, Result.Data4(3)) Then
        Exit Function
    End If
    If Not TryParseHexByte(Chars, Index, Result.Data4(4)) Then
        Exit Function
    End If
    If Not TryParseHexByte(Chars, Index, Result.Data4(5)) Then
        Exit Function
    End If
    If Not TryParseHexByte(Chars, Index, Result.Data4(6)) Then
        Exit Function
    End If
    If Not TryParseHexByte(Chars, Index, Result.Data4(7)) Then
        Exit Function
    End If
    
    If Options And Braces Then
        If Chars(Index) <> corCloseBraceChar Then
            FormatError
            Exit Function
        End If
    ElseIf Options And Parens Then
        If Chars(Index) <> corCloseParenChar Then
            FormatError
            Exit Function
        End If
    End If
    
    TryParseCore = True
End Function

Private Function TryParseDash(ByRef Chars() As Integer, ByRef Index As Long) As Boolean
    If Chars(Index) = corHyphenChar Then
        Index = Index + 1
        TryParseDash = True
    Else
        FormatError
    End If
End Function

Private Function TryParseHexByte(ByRef Chars() As Integer, ByRef Index As Long, ByRef Result As Byte) As Boolean
    Dim Value As Long
    If TryParseHex(Chars, Index, 2, Value) Then
        Result = Value
        TryParseHexByte = True
    End If
End Function

Private Function TryParseHexInt16(ByRef Chars() As Integer, ByRef Index As Long, ByRef Result As Integer) As Boolean
    Dim Output As Long
    If TryParseHex(Chars, Index, 4, Output) Then
        Result = AsWord(Output)
        TryParseHexInt16 = True
    End If
End Function

Private Function TryParseHex(ByRef Chars() As Integer, ByRef Index As Long, ByVal Count As Long, ByRef Result As Long) As Boolean
    Result = 0
    
    Dim i As Long
    For i = 0 To Count - 1
        Result = Helper.ShiftLeft(Result, 4)
        
        Dim Value As Long
        If Not TryParseHexChar(Chars(i + Index), Value) Then
            Exit Function
        End If
        
        Result = Result Or Value
    Next
    
    Index = Index + Count
    TryParseHex = True
End Function

Private Function TryParseHexChar(ByVal Ch As Integer, ByRef Result As Long) As Boolean
    Select Case Ch
        Case corZeroChar To corNineChar
            Result = Ch - corZeroChar
            TryParseHexChar = True
        Case corLowerAChar To corLowerFChar
            Result = Ch - corLowerAChar + 10
            TryParseHexChar = True
        Case corUpperAChar To corUpperFChar
            Result = Ch - corUpperAChar + 10
            TryParseHexChar = True
        Case Else
            FormatError
    End Select
End Function

Private Sub FormatError()
    Set mException = Cor.NewFormatException(Environment.GetResourceString(Format_GuidUnrecognized))
End Sub

Private Sub FormatSpecificationError()
    Set mException = Cor.NewFormatException(Environment.GetResourceString(Format_InvalidGuidFormatSpecification))
End Sub






















