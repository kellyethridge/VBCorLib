Attribute VB_Name = "modBufferHelper"
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
' Module: modBufferHelper
'

''
' Helps initialize WordBuffer structures to point to an array of Integers.
' And when the WordBuffer is destroyed by VB, the WordBuffer_Release function
' is called, allowing the WordBuffer to disconnect the Array from the SafeArray structure.
'
Option Explicit

Public Type WordBuffer
    pVTable     As Long
    This        As IUnknown
    pRelease    As Long
    Data()      As Integer
    SA          As SafeArray1d
End Type

Private mpRelease As Long

''
' Initialize a WordBuffer.
'
' @param Buffer The WordBuffer to set up and connect the Array to the SafeArray structure.
' @param pData A pointer to the Integer elements of the array.
' @param Length The number of elements in the array.
'
Public Sub InitWordBuffer(ByRef Buffer As WordBuffer, ByVal pData As Long, ByVal Length As Long)
    If mpRelease = 0 Then mpRelease = FuncAddr(AddressOf WordBuffer_Release)
    With Buffer.SA
        .cbElements = 2
        .cDims = 1
        .cElements = Length
        .pvData = pData
    End With
    With Buffer
        .pVTable = VarPtr(.pVTable)
        .pRelease = mpRelease
        SAPtr(.Data) = VarPtr(.SA)
        ObjectPtr(.This) = VarPtr(.pVTable)
    End With
End Sub

''
' Called when a WordBuffer goes out of scope.
'
' @param this The WordBuffer going out of scope.
' @remarks When the WordBuffer goes out of the scope, the "this"
' object in the WordBuffer is set to Nothing, causing this function
' to be called, allowing the WordBuffer to disconnect the array.
'
Private Function WordBuffer_Release(ByRef This As WordBuffer) As Long
    SAPtr(This.Data) = vbNullPtr
    This.SA.pvData = vbNullPtr
End Function

