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
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: SingleBlockStream
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
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult

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

Private Property Let Stream_Position(ByVal RHS As Currency)

End Property

Private Property Get Stream_Position() As Currency

End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long

End Function

Private Function Stream_ReadByte() As Long

End Function

Private Property Let Stream_ReadTimeout(ByVal RHS As Long)

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
    If Count > 0 Then Call CopyMemory(mBuffer(0), Buffer(Offset + Count - mBlockSize), mBlockSize)
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)

End Sub

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)

End Property

Private Property Get Stream_WriteTimeout() As Long

End Property
