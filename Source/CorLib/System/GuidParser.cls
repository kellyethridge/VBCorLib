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

Private Const LengthOfEnclosed      As Long = 38
Private Const LengthOfDashesOnly    As Long = 36
Private Const LengthOfDigitsOnly    As Long = 32

Private Enum ParseOptions
    None = 0
    Dashes = 1
    Braces = 2
    Parens = 4
End Enum

Private mException  As Exception
Private mChars()    As Integer
Private mIndex      As Long


Friend Sub ThrowError()
    Throw mException
End Sub

Friend Function TryParse(ByRef InputString As String, ByRef Result As VBGUID) As Boolean
    Dim Trimmed As String
    
    Trimmed = CorString.Trim(InputString)
    Init Trimmed
    
    Select Case Len(Trimmed)
        Case LengthOfEnclosed
            If AscW(Trimmed) = vbOpenBraceChar Then
                TryParse = TryParseCore(Braces Or Dashes, Result)
            Else
                TryParse = TryParseCore(Parens Or Dashes, Result)
            End If
        Case LengthOfDashesOnly
            TryParse = TryParseCore(Dashes, Result)
        Case LengthOfDigitsOnly
            TryParse = TryParseCore(ParseOptions.None, Result)
        Case Else
            TryParse = TryParseHexCore(Result)
    End Select
    
    FreeChars mChars
End Function

Friend Function TryParseExact(ByRef InputString As String, ByVal Format As Integer, ByRef Result As VBGUID) As Boolean
    Dim Trimmed As String
    
    Trimmed = CorString.Trim(InputString)
    Init Trimmed
    
    Select Case Format
        Case vbLowerBChar, vbUpperBChar
            If Len(Trimmed) = LengthOfEnclosed Then
                TryParseExact = TryParseCore(Dashes Or Braces, Result)
            Else
                FormatError
            End If
        Case vbLowerPChar, vbUpperPChar
            If Len(Trimmed) = LengthOfEnclosed Then
                TryParseExact = TryParseCore(Dashes Or Parens, Result)
            Else
                FormatError
            End If
        Case vbLowerNChar, vbUpperNChar
            If Len(Trimmed) = LengthOfDigitsOnly Then
                TryParseExact = TryParseCore(ParseOptions.None, Result)
            Else
                FormatError
            End If
        Case vbLowerDChar, vbUpperDChar
            If Len(Trimmed) = LengthOfDashesOnly Then
                TryParseExact = TryParseCore(Dashes, Result)
            Else
                FormatError
            End If
        Case vbLowerXChar, vbUpperXChar
            TryParseExact = TryParseHexCore(Result)
        Case Else
            FormatSpecificationError
    End Select
    
    FreeChars mChars
End Function

Private Sub Init(ByRef InputString As String)
    mChars = AllocChars(InputString)
    mIndex = 0
End Sub

Private Function EatAllWhiteSpace() As Long
    Dim Length      As Long
    Dim SrcIndex    As Long
    Dim DstIndex    As Long
    
    Length = UBound(mChars) + 1
    Do While SrcIndex < Length
        If Not IsWhiteSpace(mChars(SrcIndex)) Then
            mChars(DstIndex) = mChars(SrcIndex)
            DstIndex = DstIndex + 1
        End If
        SrcIndex = SrcIndex + 1
    Loop
    
    EatAllWhiteSpace = DstIndex
End Function

Private Function TryParseHexCore(ByRef Result As VBGUID) As Boolean
    If EatAllWhiteSpace <> 68 Then
        FormatError
        Exit Function
    End If

    If Not TryConsumeChar(vbOpenBraceChar) Then
        Exit Function
    End If
    
    If Not TryConsumeHexPrefix(False) Then
        Exit Function
    End If
    
    If Not TryParseNumber(8, Result.Data1) Then
        Exit Function
    End If
        
    If Not TryParseHexInt16(True, Result.Data2) Then
        Exit Function
    End If
    
    If Not TryParseHexInt16(True, Result.Data3) Then
        Exit Function
    End If
    
    If Not TryConsumeChar(vbCommaChar) Then
        Exit Function
    End If
    
    If Not TryConsumeChar(vbOpenBraceChar) Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To 7
        If Not TryParseHexByte(i > 0, Result.Data4(i)) Then
            Exit Function
        End If
    Next
    
    If Not TryConsumeChar(vbCloseBraceChar) Then
        Exit Function
    End If
    
    If Not TryConsumeChar(vbCloseBraceChar) Then
        Exit Function
    End If
    
    TryParseHexCore = True
End Function

Private Function TryParseCore(ByVal Options As ParseOptions, ByRef Result As VBGUID) As Boolean
    If Options And Braces Then
        If Not TryConsumeChar(vbOpenBraceChar) Then
            Exit Function
        End If
    ElseIf Options And Parens Then
        If Not TryConsumeChar(vbOpenParenChar) Then
            Exit Function
        End If
    End If
    
    If Not TryParseNumber(8, Result.Data1) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryConsumeChar(vbHyphenChar) Then
            Exit Function
        End If
    End If
    
    If Not TryParseInt16(Result.Data2) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryConsumeChar(vbHyphenChar) Then
            Exit Function
        End If
    End If
    
    If Not TryParseInt16(Result.Data3) Then
        Exit Function
    End If
    
    If Options And Dashes Then
        If Not TryConsumeChar(vbHyphenChar) Then
            Exit Function
        End If
    End If
    
    Dim i As Long
    For i = 0 To 7
        If i = 2 Then
            If Options And Dashes Then
                If Not TryConsumeChar(vbHyphenChar) Then
                    Exit Function
                End If
            End If
        End If
        
        If Not TryParseByte(Result.Data4(i)) Then
            Exit Function
        End If
    Next

    If Options And Braces Then
        If Not TryConsumeChar(vbCloseBraceChar) Then
            Exit Function
        End If
    ElseIf Options And Parens Then
        If Not TryConsumeChar(vbCloseParenChar) Then
            Exit Function
        End If
    End If
    
    TryParseCore = True
End Function

Private Function TryConsumeChar(ByVal Ch As Integer) As Boolean
    If mChars(mIndex) = Ch Then
        mIndex = mIndex + 1
        TryConsumeChar = True
    Else
        FormatError
    End If
End Function

Private Function TryConsumeHexPrefix(ByVal ConsumeComma As Boolean) As Boolean
    If ConsumeComma Then
        If Not TryConsumeChar(vbCommaChar) Then
            Exit Function
        End If
    End If
    
    If Not TryConsumeChar(vbZeroChar) Then
        Exit Function
    End If
    
    Dim NextChar As Integer
    NextChar = mChars(mIndex)
    If (NextChar <> vbLowerXChar) And (NextChar <> vbUpperXChar) Then
        FormatError
        Exit Function
    End If
    
    mIndex = mIndex + 1
    TryConsumeHexPrefix = True
End Function

Private Function TryParseHexByte(ByVal ConsumeComma As Boolean, ByRef Result As Byte) As Boolean
    If Not TryConsumeHexPrefix(ConsumeComma) Then
        Exit Function
    End If

    TryParseHexByte = TryParseByte(Result)
End Function

Private Function TryParseByte(ByRef Result As Byte) As Boolean
    Dim Output As Long
        
    If TryParseNumber(2, Output) Then
        Result = Output
        TryParseByte = True
    End If
End Function

Private Function TryParseHexInt16(ByVal ConsumeComma As Boolean, ByRef Result As Integer) As Boolean
    If Not TryConsumeHexPrefix(ConsumeComma) Then
        Exit Function
    End If
    
    TryParseHexInt16 = TryParseInt16(Result)
End Function

Private Function TryParseInt16(ByRef Result As Integer) As Boolean
    Dim Output As Long
    
    If TryParseNumber(4, Output) Then
        Result = AsWord(Output)
        TryParseInt16 = True
    End If
End Function

Private Function TryParseNumber(ByVal Count As Long, ByRef Result As Long) As Boolean
    Result = 0
    
    Dim i As Long
    For i = 0 To Count - 1
        Result = Helper.ShiftLeft(Result, 4)
        
        Dim Value As Long
        If Not TryParseHexChar(mChars(i + mIndex), Value) Then
            Exit Function
        End If
        
        Result = Result Or Value
    Next
    
    mIndex = mIndex + Count
    TryParseNumber = True
End Function

Private Function TryParseHexChar(ByVal Ch As Integer, ByRef Result As Long) As Boolean
    Select Case Ch
        Case vbZeroChar To vbNineChar
            Result = Ch - vbZeroChar
            TryParseHexChar = True
        Case vbLowerAChar To vbLowerFChar
            Result = Ch - vbLowerAChar + 10
            TryParseHexChar = True
        Case vbUpperAChar To vbUpperFChar
            Result = Ch - vbUpperAChar + 10
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






















