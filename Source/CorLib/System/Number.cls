VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberStatic"
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
' Module: Number
'
Option Explicit

Private Type NumberBuffer
    Scale       As Long
    Precision   As Long
    IsNegative  As Boolean
    Digits()    As Byte
    DigitsSA    As SafeArray1d
End Type

Private Enum FormatSpecifier
    UserDefinedSpecifier = 0
    GeneralUpperSpecifier = 71      ' G
    GeneralLowerSpecifier = 103     ' g
    DecimalUpperSpecifier = 68      ' D
    DecimalLowerSpecifier = 100     ' d
    NumberUpperSpecifier = 78       ' N
    NumberLowerSpecifier = 110      ' n
    HexUpperSpecifier = 88          ' X
    HexLowerSpecifier = 120         ' x
    ExponentUpperSpecifier = 69     ' E
    ExponentLowerSpecifier = 101    ' e
    FixedUpperSpecifier = 70        ' F
    FixedLowerSpecifier = 102       ' f
    CurrencyUpperSpecifier = 67     ' C
    CurrencyLowerSpecifier = 99     ' c
    PercentUpperSpecifier = 80      ' P
    PercentLowerSpecifier = 112     ' p
End Enum

Private Type NumberSpecifier
    Format      As FormatSpecifier
    Precision   As Long
End Type

Friend Function FormatInt32(ByVal Value As Long, ByRef Format As String, ByRef Info As NumberFormatInfo) As String
    Debug.Assert Not Info Is Nothing
    
    Dim Specifier As NumberSpecifier
    Specifier = ParseSpecifiers(Format, GeneralUpperSpecifier, 10)
    
End Function


Private Function ParseSpecifiers(ByRef Format As String, ByVal DefaultFormat As FormatSpecifier, ByVal DefaultPrecision As Long) As NumberSpecifier
    Select Case Len(Format)
        Case 0
            ParseSpecifiers.Format = DefaultFormat
            ParseSpecifiers.Precision = DefaultPrecision
            
        Case Is > 3
            ParseSpecifiers.Format = UserDefinedSpecifier
            ParseSpecifiers.Precision = DefaultPrecision
            
        Case Else
            Dim Chars() As Integer
            Chars = AllocChars(Format)
            
            Select Case Chars(0)
                Case corLowerAChar To corLowerZChar, corUpperAChar To corUpperZChar
                    ParseSpecifiers.Format = Chars(0)
                    
                    Dim Precision   As Long
                    Dim Ch          As Long
                    Dim i           As Long
                    For i = 1 To UBound(Chars)
                        Ch = Chars(i)
                        
                        Select Case Ch
                            Case corZeroChar To corNineChar
                                Precision = Precision * 10 + Ch - corZeroChar
                            Case Else
                                ParseSpecifiers.Format = UserDefinedSpecifier
                                ParseSpecifiers.Precision = DefaultPrecision
                                GoTo Finally
                        End Select
                    Next i
                    
                    ParseSpecifiers.Precision = Precision
                Case Else
                    ParseSpecifiers.Format = UserDefinedSpecifier
                    ParseSpecifiers.Precision = DefaultPrecision
            End Select
            
Finally:
            FreeChars Chars
    End Select
End Function

















