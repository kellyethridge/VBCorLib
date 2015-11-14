VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements Stream

Public CanWrite     As Boolean
Public CanRead      As Boolean
Public CanSeek      As Boolean
Public FlushCalled  As Boolean
Public IsClosed     As Boolean


Private Sub Class_Initialize()
    CanWrite = True
    CanRead = True
    CanSeek = True
End Sub

Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult

End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult

End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = CanRead
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = CanSeek
End Property

Private Property Get Stream_CanTimeout() As Boolean

End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = CanWrite
End Property

Private Sub Stream_CloseStream()
    IsClosed = True
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long

End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)

End Sub

Private Function Stream_Equals(Value As Variant) As Boolean

End Function

Private Sub Stream_Flush()
    FlushCalled = True
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

End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)

End Sub

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)

End Property

Private Property Get Stream_WriteTimeout() As Long

End Property
