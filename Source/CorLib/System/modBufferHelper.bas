Attribute VB_Name = "modBufferHelper"
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modBufferHelper
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

