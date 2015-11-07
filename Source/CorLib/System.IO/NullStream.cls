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
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: NullStream
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
'   IObject Interface
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
'   Stream Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim Ret As StreamAsyncResult
    
    Set Ret = New StreamAsyncResult
    With Ret
        .IsCompleted = True
        .CompletedSynchronously = True
        .IsReadType = True
    End With
    If Not Callback Is Nothing Then Call Callback.Execute(Ret)
    Set Stream_BeginRead = Ret
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim Ret As StreamAsyncResult
    
    Set Ret = New StreamAsyncResult
    With Ret
        .IsCompleted = True
        .CompletedSynchronously = True
    End With
    If Not Callback Is Nothing Then Call Callback.Execute(Ret)
    Set Stream_BeginWrite = Ret
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

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Dim Result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set Result = AsyncResult
    If Not Result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If Result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")
    
    If Not Result.Exception Is Nothing Then Throw Result.Exception
    
    Stream_EndRead = Result.BytesRead
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    Dim Result As StreamAsyncResult
    
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    Set Result = AsyncResult
    If Result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If Result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")

    If Not Result.Exception Is Nothing Then Throw Result.Exception
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
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Timeouts))
End Property

Private Property Get Stream_ReadTimeout() As Long
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Timeouts))
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
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Timeouts))
End Property

Private Property Get Stream_WriteTimeout() As Long
    Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_Timeouts))
End Property
