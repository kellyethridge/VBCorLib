Attribute VB_Name = "Startup"
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: modMain
'
'@Folder("CorLib.System")
Option Explicit

Private mInIDE        As Boolean
Private mInDebugger   As Boolean

Public Property Get InIDE() As Boolean
    InIDE = mInIDE
End Property

Public Property Get InDebugger() As Boolean
    InDebugger = mInDebugger
End Property

Private Sub Main()
    InitHelper
    SetInIDE
    SetInDebugger
    InitMissing
    InitMathematics
    InitGlobalization
    InitEncoding
End Sub

Private Sub InitMissing(Optional ByRef Value As Variant)
    Missing = Value
End Sub

''
' This is to determine if the compiled dll is being used in the
' VB6 IDE by another project. This is primarily so that the console
' class can disable the exit button if we are running in an IDE.
'
Private Sub SetInDebugger()
    Const BufferSize As Long = 512
    Dim Result As String
    Dim Length As Long
    
    Result = String$(BufferSize, 0)
    Length = GetModuleFileNameW(vbNullPtr, Result, BufferSize)
    Result = Left$(Result, Length)
    mInDebugger = (UCase$(Right$(Result, 8)) = "\VB6.EXE")
End Sub

Private Sub SetInIDE()
    On Error GoTo errTrap
    Debug.Assert 1 \ 0
    Exit Sub
    
errTrap:
    mInIDE = True
End Sub
