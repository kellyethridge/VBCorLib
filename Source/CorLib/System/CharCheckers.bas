Attribute VB_Name = "CharCheckers"
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: CharCheckers
'
Option Explicit

Private Const vbLeftCheveronChar    As Long = 60
Private Const vbRightCheveronChar   As Long = 62
Private Const vbSpaceChar           As Long = 32
Private Const vbQuoteChar           As Long = 34


Public Function IsValidTagChar(ByRef Char As Integer) As Boolean
    Select Case Char
        Case vbLeftCheveronChar, vbRightCheveronChar, vbSpaceChar
            Exit Function
    End Select
    
    IsValidTagChar = True
End Function

Public Function IsValidTextChar(ByRef Char As Integer) As Boolean
    Select Case Char
        Case vbLeftCheveronChar, vbRightCheveronChar
            Exit Function
    End Select
    
    IsValidTextChar = True
End Function

Public Function IsValidAttributeValueChar(ByRef Char As Integer) As Boolean
    Select Case Char
        Case vbLeftCheveronChar, vbRightCheveronChar, vbQuoteChar
            Exit Function
    End Select
    
    IsValidAttributeValueChar = True
End Function

Public Function IsWhiteSpaceChar(ByRef Char As Integer) As Boolean
    IsWhiteSpaceChar = Statics.Char.IsWhiteSpaceChar(Char)
End Function
