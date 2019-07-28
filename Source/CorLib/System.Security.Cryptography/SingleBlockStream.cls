VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleBlockStream"
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
' Module: SingleBlockStream
'

''
' This provides a small stream buffer that buffers only the last block of a write process.
'
Option Explicit
Implements Stream

Private mBuffer() As Byte
Private mBlockSize As Long


''
' Returns the last block buffered.
'
Friend Property Get Hash() As Byte()
    Hash = mBuffer
End Property

''
' Inits the size of the block to buffer.
Friend Sub Init(ByVal BlockSize As Long)
    mBlockSize = BlockSize
    ReDim mBuffer(0 To BlockSize - 1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Stream Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult

End Function

Private Property Get Stream_CanRead() As Boolean

End Property

Private Property Get Stream_CanSeek() As Boolean

End Property

Private Property Get Stream_CanTimeout() As Boolean

End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = True
End Property

Private Sub Stream_CloseStream()

End Sub

Private Sub Stream_CopyTo(ByVal Destination As Stream, Optional ByVal BufferSize As Long = 81920)

End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long

End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)

End Sub

Private Function Stream_Equals(Value As Variant) As Boolean

End Function

Private Sub Stream_Flush()

End Sub

Private Function Stream_GetHashCode() As Long

End Function

Private Property Get Stream_Length() As Currency

End Property

Private Property Let Stream_Position(ByVal Value As Currency)

End Property

Private Property Get Stream_Position() As Currency

End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long

End Function

Private Function Stream_ReadByte() As Long

End Function

Private Property Let Stream_ReadTimeout(ByVal Value As Long)

End Property

Private Property Get Stream_ReadTimeout() As Long

End Property

Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency

End Function

Private Sub Stream_SetLength(ByVal Value As Currency)

End Sub

Private Function Stream_ToString() As String

End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    ' Since this is for internal use, we don't expect any partial blocks.
    If Count > 0 Then
        CopyMemory mBuffer(0), Buffer(Offset + Count - mBlockSize), mBlockSize
    End If
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)

End Sub

Private Property Let Stream_WriteTimeout(ByVal Value As Long)

End Property

Private Property Get Stream_WriteTimeout() As Long

End Property
