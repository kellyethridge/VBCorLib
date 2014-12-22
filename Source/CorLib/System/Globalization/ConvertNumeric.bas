Attribute VB_Name = "ConvertNumeric"
'The MIT License (MIT)
'Copyright (c) 2013 Kelly Ethridge
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
' Module: ConvertNumeric
'
Option Explicit

Private Const UnknownPrecision As Long = -1

Private Enum FormatSpecifier
    CustomSpecifier = -1
    GeneralSpecifier = 71      ' G
    DecimalSpecifier = 68      ' D
    NumberSpecifier = 78       ' N
    HexSpecifier = 88          ' X
    ExponentSpecifier = 69     ' E
    FixedSpecifier = 70        ' F
    CurrencySpecifier = 67     ' C
    PercentSpecifier = 80      ' P
End Enum

Private mFormatSafeArray As SafeArray1d


Public Function ToStringFormat(ByRef FormatInfo As NumberFormatInfo, ByRef Value As Variant, ByRef Format As String) As String
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
        
        Case Else
            ToStringFormat = Value
    End Select
End Function



Private Sub ParseFormat(ByRef Format As String, ByRef FormatSpecifier As Long, ByRef PrecisionSpecifier As Long)
    Dim FormatLength As Long
    FormatLength = Len(Format)
    
    Select Case FormatLength
        Case 0
            FormatSpecifier = GeneralSpecifier
            PrecisionSpecifier = UnknownPrecision
            
        Case Is > 3
            FormatSpecifier = CustomSpecifier
            PrecisionSpecifier = UnknownPrecision
            
        Case Else
            Dim Chars() As Integer
            AttachCharsQuick Format, Chars
            
            FormatSpecifier = Chars(0)
            
            Select Case FormatSpecifier
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    If FormatLength = 1 Then
                        PrecisionSpecifier = UnknownPrecision
                        DetachCharsQuick Chars
                        Exit Sub
                    End If
                    
                    PrecisionSpecifier = 0
                    
                    Dim ch As Long
                    Dim i As Long
                    For i = 1 To FormatLength - 1
                        ch = Chars(i)
                        
                        Select Case ch
                            Case vbZero To vbNineChar
                                PrecisionSpecifier = PrecisionSpecifier * 10 + ch - vbZero
                            Case Else
                                FormatSpecifier = CustomSpecifier
                                PrecisionSpecifier = UnknownPrecision
                                DetachCharsQuick Chars
                                Exit Sub
                        End Select
                    Next i
                    
                Case Else
                    FormatSpecifier = CustomSpecifier
                    PrecisionSpecifier = UnknownPrecision
                    DetachCharsQuick Chars
            End Select
    End Select
End Sub
