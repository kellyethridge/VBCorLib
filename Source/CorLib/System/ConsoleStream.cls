VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsoleStream"
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
' Module: ConsoleStream
'

''
' This is a <b>Stream</b> wrapper around the console input and output methods.
'
'@Folder("CorLib.System")
Option Explicit
Implements IObject
Implements Stream

Private mHandle         As Long
Private mCanRead        As Boolean
Private mCanWrite       As Boolean
Private mSingleByte(0)  As Byte



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Handle As Long, ByVal Access As FileAccess)
    mHandle = Handle
    mCanRead = CBool(Access And ReadAccess)
    mCanWrite = CBool(Access And WriteAccess)
End Sub


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
    Dim Ret As StreamAsyncResult
    
    Set Ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo Catch
    
    With Ret
        .BytesRead = Stream_ReadBlock(Buffer, Offset, Count)
        .IsCompleted = True
        .IsReadType = True
    End With
    
    If Not Callback Is Nothing Then
        Callback.Execute Ret
    End If
    
Catch:
    Dim Ex As Exception
    If Catch(Ex, Err) Then
        Set Ret.Exception = Ex
    End If
    
    Set Stream_BeginRead = Ret
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Ret As StreamAsyncResult
    
    Set Ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo Catch
    
    With Ret
        .CompletedSynchronously = True
        .IsCompleted = True
    End With
    
    Stream_WriteBlock Buffer, Offset, Count
    
    If Not Callback Is Nothing Then
        Callback.Execute Ret
    End If
    
Catch:
    Dim Ex As Exception
    If Catch(Ex, Err) Then
        Set Ret.Exception = Ex
    End If
    
    Set Stream_BeginWrite = Ret
End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = True
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = False
End Property

Private Property Get Stream_CanTimeout() As Boolean
    Stream_CanTimeout = False
End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = True
End Property

Private Sub Stream_CloseStream()
    mCanRead = False
    mCanWrite = False
End Sub

Private Sub Stream_CopyTo(ByVal Destination As Stream, Optional ByVal BufferSize As Long = 81920)
    Statics.Stream.CopyTo Me, Destination, BufferSize
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Dim Result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult", "AsyncResult object is required.")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    
    Set Result = AsyncResult
    
    If Not Result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If Result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")
    If Not Result.Exception Is Nothing Then _
        Throw Result.Exception
    
    Stream_EndRead = Result.BytesRead
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    Dim Result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult", "AsyncResult object is required.")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set Result = AsyncResult
    If Result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If Result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")

    If Not Result.Exception Is Nothing Then _
        Throw Result.Exception
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
    Throw New NotSupportedException
End Property

Private Property Let Stream_Position(ByVal RHS As Currency)
    Throw New NotSupportedException
End Property

Private Property Get Stream_Position() As Currency
    Throw New NotSupportedException
End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    ValidateByteRange Buffer, Offset, Count, NameOfBuffer, NameOfOffset
    
    If Not mCanRead Then _
        Throw Cor.NewNotSupportedException("Stream does not support reading")
    
    If ReadFile(mHandle, Buffer(Offset), Count, Stream_ReadBlock, ByVal 0&) = BOOL_FALSE Then _
        Statics.Error.Win32Error Err.LastDllError
End Function

Private Function Stream_ReadByte() As Long
    If Stream_ReadBlock(mSingleByte, 0, 1) > 0 Then
        Stream_ReadByte = mSingleByte(0)
    Else
        Stream_ReadByte = -1
    End If
End Function

Private Property Let Stream_ReadTimeout(ByVal RHS As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

Private Property Get Stream_ReadTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    Throw New NotSupportedException
End Function

Private Sub Stream_SetLength(ByVal Value As Currency)
    Throw New NotSupportedException
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = IObject_ToString
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    ValidateByteRange Buffer, Offset, Count, NameOfBuffer, NameOfOffset
    
    If Not mCanWrite Then _
        Throw Cor.NewNotSupportedException("The stream does not support writing.")
    
    If WriteFile(mHandle, Buffer(Offset), Count, 0, ByVal 0&) = BOOL_FALSE Then _
        Statics.Error.Win32Error Err.LastDllError
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)
    mSingleByte(0) = Value
    Stream_WriteBlock mSingleByte, 0, 1
End Sub

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

Private Property Get Stream_WriteTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property
