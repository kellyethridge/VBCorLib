VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullStream"
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
' Module: NullStream
'

''
' Provides a stream with no backing. Data can be written, but nothing will be stored.
'
' @see Stream
'
Option Explicit
Implements IObject
Implements Stream



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Stream
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Result As New StreamAsyncResult
    Result.IsCompleted = True
    Result.CompletedSynchronously = True
    Result.IsReadType = True
    
    If Not Callback Is Nothing Then
        Callback.Execute Result
    End If
    
    Set Stream_BeginRead = Result
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Result As New StreamAsyncResult
    Result.IsCompleted = True
    Result.CompletedSynchronously = True

    If Not Callback Is Nothing Then
        Callback.Execute Result
    End If
    
    Set Stream_BeginWrite = Result
End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = True
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = True
End Property

Private Property Get Stream_CanTimeout() As Boolean
    Stream_CanTimeout = False
End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = True
End Property

Private Sub Stream_CloseStream()
' do nothing
End Sub

Private Sub Stream_CopyTo(ByVal Destination As Stream, Optional ByVal BufferSize As Long = 81920)
' do nothing
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Dim Result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Error.ArgumentNull "AsyncResult"
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    Set Result = AsyncResult
    If Not Result.IsReadType Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    If Result.EndCalled Then _
        Error.InvalidOperation InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    
    If Not Result.Exception Is Nothing Then
        Throw Result.Exception
    End If
    
    Stream_EndRead = Result.BytesRead
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    Dim Result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Error.ArgumentNull "AsyncResult"
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple
    
    Set Result = AsyncResult
    If Result.IsReadType Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple
    If Result.EndCalled Then _
        Error.InvalidOperation InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple

    If Not Result.Exception Is Nothing Then
        Throw Result.Exception
    End If
End Sub

Private Function Stream_Equals(Value As Variant) As Boolean
    Stream_Equals = IObject_Equals(Value)
End Function

Private Sub Stream_Flush()
' do nothing
End Sub

Private Function Stream_GetHashCode() As Long
    Stream_GetHashCode = IObject_GetHashCode
End Function

Private Property Get Stream_Length() As Currency
' do nothing
End Property

Private Property Let Stream_Position(ByVal RHS As Currency)
' do nothing
End Property

Private Property Get Stream_Position() As Currency
' do nothing
End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
' do nothing
End Function

Private Function Stream_ReadByte() As Long
    Stream_ReadByte = -1
End Function

Private Property Let Stream_ReadTimeout(ByVal RHS As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

Private Property Get Stream_ReadTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
' do nothing
End Function

Private Sub Stream_SetLength(ByVal Value As Currency)
' do nothing
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = IObject_ToString
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
' do nothing
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)
' do nothing
End Sub

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

Private Property Get Stream_WriteTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property
