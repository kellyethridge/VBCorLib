Attribute VB_Name = "ConsoleCallbacks"
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
' Module: ConsoleCallbacks
'
'@Folder("CorLib.System")
Option Explicit

''
' We keep these flags here instead of the Console class because the
' ControlBreakHandler callback routine is not called in a threadsafe
' manor. We don't want to be crossing threads when calling into a
' COM object and potentially corrupt memory causing a crash.
Private mBreak      As Boolean
Private mBreakType  As ConsoleBreakType


Public Property Get Break() As Boolean
    Break = mBreak
End Property

Public Property Let Break(ByVal Value As Boolean)
    mBreak = Value
End Property

Public Property Get BreakType() As ConsoleBreakType
    BreakType = mBreakType
End Property

''
' This is the callback used by the SetConsoleCtrlHandler API.
' This function is not called in a threadsafe manor, so don't
' use any COM objects during the routine.
Public Function ControlBreakHandler(ByVal dwCtrlType As Long) As Long
    mBreak = True
    mBreakType = dwCtrlType
    ControlBreakHandler = True
End Function
