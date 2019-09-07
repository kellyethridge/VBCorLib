Attribute VB_Name = "CharCallbacks"
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

''
' These are used for callbacks for such functions a CorString.AllChars().
'
Option Explicit

Public Function IsValidTagCallback(ByRef Char As Integer) As Boolean
    Select Case Char
        Case vbLeftCheveronChar, vbRightCheveronChar, vbSpaceChar
            Exit Function
    End Select
    
    IsValidTagCallback = True
End Function

Public Function IsValidTextCallback(ByRef Char As Integer) As Boolean
    Select Case Char
        Case vbLeftCheveronChar, vbRightCheveronChar
            Exit Function
    End Select
    
    IsValidTextCallback = True
End Function

Public Function IsValidAttributeValueCallback(ByRef Char As Integer) As Boolean
    Select Case Char
        Case vbLeftCheveronChar, vbRightCheveronChar, vbQuoteChar
            Exit Function
    End Select
    
    IsValidAttributeValueCallback = True
End Function

Public Function IsPeriodCallback(ByRef Char As Integer) As Boolean
    IsPeriodCallback = Char = vbPeriodChar
End Function
