VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: StreamStatic
'

''
' Provides static methods for the <b>Stream</b> class.
'
' @see Stream
' @see PublicFunctions
'
Option Explicit

Private mNullStream As New NullStream

''
' Returns a Stream object with no backing store.
'
' @return A Stream object that has no backing store.
'
Public Property Get NullStream() As Stream
    Set NullStream = mNullStream
End Property

Friend Sub CopyTo(ByVal Source As Stream, ByVal Destination As Stream, ByVal BufferSize As Long)
    Dim Buffer() As Byte
    Dim ReadCount As Long
    
    If Source Is Nothing Then _
        Error.ArgumentNull "Source"
    If Destination Is Nothing Then _
        Error.ArgumentNull "Destination"
    If BufferSize < 1 Then _
        Error.ArgumentOutOfRange "BufferSize", ArgumentOutOfRange_NeedPosNum
    If Not Source.CanRead And Not Source.CanWrite Then _
        Throw Cor.NewObjectDisposedException("Source", Environment.GetResourceString(ObjectDisposed_StreamClosed))
    If Not Destination.CanRead And Not Destination.CanWrite Then _
        Throw Cor.NewObjectDisposedException("Destination", Environment.GetResourceString(ObjectDisposed_StreamClosed))
    If Not Source.CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
    If Not Destination.CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    
    ReDim Buffer(0 To BufferSize - 1)
    
    Do
        ReadCount = Source.ReadBlock(Buffer, 0, BufferSize)
        
        If ReadCount > 0 Then
            Destination.WriteBlock Buffer, 0, ReadCount
        End If
    Loop While ReadCount = BufferSize
End Sub
