VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoStreamSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements Stream
Implements CryptoStream

Public Mode As CryptoStreamMode
Public FlushFinalBlockCount As New Count


Private Function CryptoStream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As CorLib.AsyncCallback, Optional State As Variant) As CorLib.IAsyncResult

End Function

Private Function CryptoStream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As CorLib.AsyncCallback, Optional State As Variant) As CorLib.IAsyncResult

End Function

Private Property Get CryptoStream_CanRead() As Boolean
    CryptoStream_CanRead = (Mode = ReadMode)
End Property

Private Property Get CryptoStream_CanSeek() As Boolean

End Property

Private Property Get CryptoStream_CanTimeout() As Boolean

End Property

Private Property Get CryptoStream_CanWrite() As Boolean
    CryptoStream_CanWrite = (Mode = WriteMode)
End Property

Private Sub CryptoStream_CloseStream()

End Sub

Private Sub CryptoStream_CopyTo(ByVal Destination As CorLib.Stream, Optional ByVal BufferSize As Long = 81920)

End Sub

Private Function CryptoStream_EndRead(ByVal AsyncResult As CorLib.IAsyncResult) As Long

End Function

Private Sub CryptoStream_EndWrite(ByVal AsyncResult As CorLib.IAsyncResult)

End Sub

Private Function CryptoStream_Equals(Value As Variant) As Boolean

End Function

Private Sub CryptoStream_Flush()

End Sub

Private Sub CryptoStream_FlushFinalBlock()
    FlushFinalBlockCount.Increment
End Sub

Private Function CryptoStream_GetHashCode() As Long

End Function

Private Property Get CryptoStream_Length() As Currency

End Property

Private Property Get CryptoStream_Position() As Currency

End Property

Private Property Let CryptoStream_Position(ByVal RHS As Currency)

End Property

Private Function CryptoStream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long

End Function

Private Function CryptoStream_ReadByte() As Long

End Function

Private Property Get CryptoStream_ReadTimeout() As Long

End Property

Private Property Let CryptoStream_ReadTimeout(ByVal RHS As Long)

End Property

Private Function CryptoStream_SeekPosition(ByVal Offset As Currency, ByVal Origin As CorLib.SeekOrigin) As Currency

End Function

Private Sub CryptoStream_SetLength(ByVal Value As Currency)

End Sub

Private Function CryptoStream_ToString() As String

End Function

Private Sub CryptoStream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)

End Sub

Private Sub CryptoStream_WriteByte(ByVal Value As Byte)

End Sub

Private Property Get CryptoStream_WriteTimeout() As Long

End Property

Private Property Let CryptoStream_WriteTimeout(ByVal RHS As Long)

End Property

Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As CorLib.AsyncCallback, Optional State As Variant) As CorLib.IAsyncResult

End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As CorLib.AsyncCallback, Optional State As Variant) As CorLib.IAsyncResult

End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = (Mode = ReadMode)
End Property

Private Property Get Stream_CanSeek() As Boolean

End Property

Private Property Get Stream_CanTimeout() As Boolean

End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = (Mode = WriteMode)
End Property

Private Sub Stream_CloseStream()

End Sub

Private Sub Stream_CopyTo(ByVal Destination As CorLib.Stream, Optional ByVal BufferSize As Long = 81920)

End Sub

Private Function Stream_EndRead(ByVal AsyncResult As CorLib.IAsyncResult) As Long

End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As CorLib.IAsyncResult)

End Sub

Private Function Stream_Equals(Value As Variant) As Boolean

End Function

Private Sub Stream_Flush()

End Sub

Private Function Stream_GetHashCode() As Long

End Function

Private Property Get Stream_Length() As Currency

End Property

Private Property Get Stream_Position() As Currency

End Property

Private Property Let Stream_Position(ByVal RHS As Currency)

End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long

End Function

Private Function Stream_ReadByte() As Long

End Function

Private Property Get Stream_ReadTimeout() As Long

End Property

Private Property Let Stream_ReadTimeout(ByVal RHS As Long)

End Property

Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As CorLib.SeekOrigin) As Currency

End Function

Private Sub Stream_SetLength(ByVal Value As Currency)

End Sub

Private Function Stream_ToString() As String

End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)

End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)

End Sub

Private Property Get Stream_WriteTimeout() As Long

End Property

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)

End Property
