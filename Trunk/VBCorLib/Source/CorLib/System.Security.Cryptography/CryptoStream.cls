VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoStream"
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
' Module: CryptoStream
'

''
' Provides a streaming mechanism for ciphering data.
'
' @remarks A <b>CryptoStream</b> can be written to or read from in the same manner
' as other stream objects. A <b>CryptoStream</b> is designed to be used in a daisy-chain
' fashion, allowing for multiple transformations to be applied as a single process.
' <p>By supplying another <b>Stream</b> object to the <b>Cor.NewCryptoStream</b> function,
' an automatic daisy-chaining of the two streams is created. The <b>Stream</b> does not
' have to be another <b>CryptoStream</b> object. It can be any object implementing the
' <b>Stream</b> interface.
'
' @see Constructors
'
Option Explicit
Implements IObject
Implements Stream



''
' Specifies modes for a CryptoStream.
'
' @param ReadMode Sets the CryptoStream to a reading mode.
' @param WriteMode Sets the CryptoStream to a writing mode.
'
Public Enum CryptoStreamMode
    ReadMode = 0
    WriteMode = 1
End Enum

Private mMode               As CryptoStreamMode
Private mStream             As Stream
Private mTransform          As ICryptoTransform
Private mInputBuffer()      As Byte
Private mOutputBuffer()     As Byte
Private mInputIndex         As Long
Private mOutputIndex        As Long
Private mBytesBuffered      As Long
Private mFinalTransformed   As Boolean
Private mOutputBlockSize    As Long
Private mInputBlockSize     As Long



''
' Returns if the Stream can be read from.
'
' @return Returns True if the Stream supports reading, False otherwise.
' @remarks A CryptoStream object supports only reading or writing, not both simultaneously.
'
Public Property Get CanRead() As Boolean
    CanRead = (mMode = ReadMode)
End Property

''
' Returns if the Stream can seek.
'
' @return Always returns False.
'
Public Property Get CanSeek() As Boolean
    CanSeek = False
End Property

''
' Returns if the Stream can be written to.
'
' @return Returns True if the Stream supports writing, False otherwise.
' @remarks A CryptoStream object supports only ready or writing, not both simultaneously.
'
Public Property Get CanWrite() As Boolean
    CanWrite = (mMode = WriteMode)
End Property

''
' Returns the current position withing the Stream.
'
' @return The current position.
' @remarks Always throws a NotSupportedException exception.
'
Public Property Get Position() As Currency
    Error.NotSupported NotSupported_UnseekableStream
End Property

''
' Sets the current position within the Stream.
'
' @param RHS The new position.
' @remarks Always throws a NotSupportedException exception.
'
Public Property Let Position(ByVal RHS As Currency)
    Error.NotSupported NotSupported_UnseekableStream
End Property

''
' Returns the length of the current Stream.
'
' @return The length of the Stream.
' @remarks Always throws a NotSupportedException exception.
'
Public Property Get Length() As Currency
    Error.NotSupported NotSupported_UnseekableStream
End Property

''
' Sets the length of the current Stream.
'
' @param Value The new length of the Stream.
' @remarks Always throws a NotSupportedException exception.
'
Public Sub SetLength(ByVal Value As Currency)
    Error.NotSupported NotSupported_UnseekableStream
End Sub

''
' Seeks a new position within the Stream.
'
' @param Offset The number of bytes to move.
' @param Origin The starting position in the stream to move from.
' @return The new position in the Stream.
' @remarks Always throws a NotSupportedException exception.
'
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    Error.NotSupported NotSupported_UnseekableStream
End Function

''
' Returns if the Stream can timeout.
'
' @returns Returns True if the Stream can timeout, False otherwise.
'
Public Property Get CanTimeout() As Boolean
    CanTimeout = mStream.CanTimeout
End Property

''
' Returns the amount of time must pass before a Read timeout occurs.
'
' @return Returns the timeout period in milliseconds.
'
Public Property Get ReadTimeout() As Long
    ReadTimeout = mStream.ReadTimeout
End Property

''
' Sets the timeout period for Read operations.
'
' @param RHS The timeout period in milliseconds.
'
Public Property Let ReadTimeout(ByVal RHS As Long)
    mStream.ReadTimeout = RHS
End Property

''
' Returns the timeout period for a write operation.
'
' @return The timeout period in milliseconds.
'
Public Property Get WriteTimeout() As Long
    WriteTimeout = mStream.WriteTimeout
End Property

''
' Sets the timeout period for a write operation.
'
' @param RHS The timeout period in milliseconds.
'
Public Property Let WriteTimeout(ByVal RHS As Long)
    mStream.WriteTimeout = RHS
End Property

''
' Closes the underlying Stream.
'
Public Sub CloseStream()
    If Me.CanWrite And (Not mFinalTransformed) Then
        FlushFinalBlock
    End If
    mStream.CloseStream
    Erase mInputBuffer
    Erase mOutputBuffer
End Sub

''
' Flushes the buffers of the underlying Stream.
'
' @remarks This does not flush the buffers within the CryptoStream. Call FlushFinalBlock
' to flush the remaining data to the underlying stream.
'
Public Sub Flush()
    mStream.Flush
End Sub

''
' Flushes the remaining data to the underlying stream.
'
' @remarks This can only be called once.
'
Public Sub FlushFinalBlock()
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    If mFinalTransformed Then _
        Throw Cor.NewNotSupportedException("Cannot call FlushFinalBlock twice.")
    
    mOutputBuffer = mTransform.TransformFinalBlock(mInputBuffer, 0, mBytesBuffered)
    mStream.WriteBlock mOutputBuffer, 0, CorArray.LengthFirstDim(mOutputBuffer)
    
    If TypeOf mStream Is CryptoStream Then
        Dim cs As CryptoStream
        Set cs = mStream
        cs.FlushFinalBlock
    End If
    
    Flush
    InitBuffers
    mFinalTransformed = True
End Sub

''
' Reads a requested amount of data from the stream.
'
' @param Buffer The byte array that will receive the requested data.
' @param Offset The starting index within <i>Buffer</i> to begin writing data.
' @param Count The number of bytes to read.
' @return The actual number of bytes read.
' @remarks The number of bytes read may be less than the actual number of bytes requested.
'
Public Function ReadBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    If Not CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
    If CorArray.IsNull(Buffer) Then _
        Error.ArgumentNull "Buffer", ArgumentNull_Array
    If Offset < LBound(Buffer) Then _
        Error.ArgumentOutOfRange "Offset", ArgumentOutOfRange_LBound
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    Dim BytesRead  As Long
    
    ' Just grab anything that might already be buffered.
    BytesRead = ReadBufferedBytes(Buffer, Offset, Count)
    
    ' Take advantage of transforming multiple blocks if possible.
    If mTransform.CanTransformMultipleBlocks Then
        BytesRead = BytesRead + ReadMultiBlock(Buffer, Offset, Count)
    End If
    
    ' Fill in the end of the requested amount a little at a time.
    Do While Count > 0
        If mBytesBuffered = 0 Then
            If Not FillBuffer Then Exit Do
        End If
        
        BytesRead = BytesRead + ReadBufferedBytes(Buffer, Offset, Count)
    Loop
    
    ReadBlock = BytesRead
End Function

''
' Reads a single byte from the Stream.
'
' @return Returns the byte data or -1 if the end of the stream was reached.
'
Public Function ReadByte() As Long
    If Not CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
    
    If mBytesBuffered = 0 Then
        If Not FillBuffer Then
            ReadByte = -1
            Exit Function
        End If
    End If
    
    ReadByte = mOutputBuffer(mOutputIndex)
    mOutputIndex = mOutputIndex + 1
    mBytesBuffered = mBytesBuffered - 1
End Function

''
' Begins an Asynchronous read operation (currently is only synchronous)
'
' @param Buffer The array to store the bytes that are read from the stream.
' @param Offset The index in <i>Buffer</i> to begin storing bytes.
' @param Count The number of bytes to be read from the stream.
' @param Callback An object that is called when the read operation is complete.
' @param State user-defined data to be kept with the <i>Callback</i> object for later retrieval.
' @return An IAsyncResult object used to identify the corrisponding Endread and report
' information about the buffer read.
'
Public Function BeginRead(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    If Not CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
    
    Dim Result As StreamAsyncResult
    Set Result = Cor.NewStreamAsyncResult(State)
    
    On Error GoTo Catch
    Result.BytesRead = ReadBlock(Buffer, Offset, Count)
    Result.IsCompleted = True
    Result.IsReadType = True

    If Not Callback Is Nothing Then
        Callback.Execute Result
    End If
    
    GoTo Finally
    
Catch:
    Dim Ex As Exception
    Catch Ex, Err
    Set Result.Exception = Ex
Finally:
    Set BeginRead = Result
End Function

''
' Signifies the end of an asynchronous read from the stream.
'
' @param AsyncResult The IAsyncResult object returned from the BeginRead function.
' @return The number of bytes read from the stream in to <i>Buffer</i> in the BeginRead method.
'
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long
    If AsyncResult Is Nothing Then _
        Error.ArgumentNull "AsyncResult", ArgumentNull_Generic
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    
    Dim Result As StreamAsyncResult
    Set Result = AsyncResult
    If Not Result.IsReadType Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    If Result.EndCalled Then _
        Error.InvalidOperation InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    
    If Not Result.Exception Is Nothing Then
        Throw Result.Exception
    End If
    
    EndRead = Result.BytesRead
End Function

''
' Writes an array of bytes to the Stream.
'
' @param Buffer The array of bytes to be written to the Stream.
' @param Offset The index within <i>Buffer</i> to begin writing from.
' @param Count The number of bytes to be written.
' @param The bytes are transformed using the supplied ICryptoTransform before
' being written to the underlying stream.
'
Public Sub WriteBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    
    ValidateArrayRange Buffer, Offset, Count, Parameter_Buffer, Parameter_Offset
    
    If (Count > mInputBlockSize) And mTransform.CanTransformMultipleBlocks Then
        WriteMultiBlock Buffer, Offset, Count
    End If
    
    Do While Count > 0
        WriteToBuffer Buffer, Offset, Count
    Loop
End Sub

''
' Writes a single byte to the Stream.
'
' @param Value The byte to be written to the stream.
'
Public Sub WriteByte(ByVal Value As Byte)
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    
    mInputBuffer(mInputIndex) = Value
    mInputIndex = mInputIndex + 1
    mBytesBuffered = mBytesBuffered + 1
    
    If mBytesBuffered = mInputBlockSize Then
        FlushBuffer
    End If
End Sub

''
' Begins an asynchronous buffer write. Currently the CryptoStream class does not
' support asynchronous buffer writing.
'
' @param Buffer The source array to copy the bytes from into the stream.
' @param Offset The starting index in the source <i>Buffer</i> to begin copying from.
' @param Count The maximum number of bytes to be copied into the stream.
' @param Callback A user supplied object to be notified of the finish of the writing.
' @param State User data that can be carried within the IAsyncResult object return.
' @return An IAsyncResult object used to identify the corrisponding EndBread and report
' information about the buffer read.
'
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Dim Result As StreamAsyncResult
    Set Result = Cor.NewStreamAsyncResult(State)
    
    On Error GoTo Catch
    
    Result.CompletedSynchronously = True
    Result.IsCompleted = True
    WriteBlock Buffer, Offset, Count
    
    If Not Callback Is Nothing Then
        Callback.Execute Result
    End If
    
    GoTo Finally
    
Catch:
    Dim Ex As Exception
    Catch Ex, Err
    Set Result.Exception = Ex
Finally:
    Set BeginWrite = Result
End Function

''
' Signifies the end of an asynchronous write to the stream.
'
' @param AsyncResult The IAsyncResult returned from the BeginWrite function.
'
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult)
    If AsyncResult Is Nothing Then _
        Error.ArgumentNull "AsyncResult", ArgumentNull_Generic
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple
    
    Dim Result As StreamAsyncResult
    Set Result = AsyncResult
    If Result.IsReadType Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple
    If Result.EndCalled Then _
        Error.InvalidOperation InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple

    If Not Result.Exception Is Nothing Then
        Throw Result.Exception
    End If
End Sub

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Stream As Stream, ByVal Transform As ICryptoTransform, ByVal Mode As CryptoStreamMode)
    If Stream Is Nothing Then _
        Error.ArgumentNull "Stream", ArgumentNull_Stream
    If Transform Is Nothing Then _
        Error.ArgumentNull "Transform", ArgumentNull_Generic
    
    Select Case Mode
        Case CryptoStreamMode.ReadMode
            If Not Stream.CanRead Then _
                Error.Argument Argument_StreamNotReadable, "Stream"
                
        Case CryptoStreamMode.WriteMode
            If Not Stream.CanWrite Then _
                Error.Argument Argument_StreamNotWritable, "Stream"
                
        Case Else
            Error.ArgumentOutOfRange "Mode", ArgumentOutOfRange_Enum
    End Select
    
    Set mStream = Stream
    Set mTransform = Transform
    mMode = Mode
    
    mOutputBlockSize = mTransform.OutputBlockSize
    mInputBlockSize = mTransform.InputBlockSize
    InitBuffers
End Sub

Private Sub Class_Terminate()
    CloseStream
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitBuffers()
    ReDim mInputBuffer(0 To mInputBlockSize - 1)
    ReDim mOutputBuffer(0 To mOutputBlockSize - 1)
End Sub

''
' This is used when reading from the Stream.
' We like to have a full buffer to read from for small reads.
'
Private Function FillBuffer() As Boolean
    Do
        ' We need to read a full block of bytes in at one time. The Stream
        ' being read from must have an exact multiple of mInputBlockSize bytes
        ' to read or it will eventually fail.
        Dim BytesRead As Long
        BytesRead = mStream.ReadBlock(mInputBuffer, 0, mInputBlockSize)
        
        ' If we did that, then we can assume we are not at the end of the Stream.
        If BytesRead > 0 Then
            ' So just transform the whole buffer to our output buffer.
            mBytesBuffered = mTransform.TransformBlock(mInputBuffer, 0, BytesRead, mOutputBuffer, 0)
        Else
            ' We didn't have any bytes, so we assume the end of the stream was reached.
            ' Using the TransformFinalBlock will strip away the padding from what we
            ' have left in our input buffer.
            mOutputBuffer = mTransform.TransformFinalBlock(mInputBuffer, 0, mBytesBuffered)
            mBytesBuffered = CorArray.Length(mOutputBuffer)
            Exit Do
        End If
    Loop While mBytesBuffered = 0
    
    mOutputIndex = 0
    FillBuffer = (mBytesBuffered > 0)
End Function

''
' This is used when writing to the Stream. It handles the small internal buffer we have.
' This will get the mInputBuffer transformed and written to the Stream.
' This assumes that the buffer is full.
'
Private Sub FlushBuffer()
    Dim TotalBytesTransformed As Long
    TotalBytesTransformed = mTransform.TransformBlock(mInputBuffer, 0, mInputBlockSize, mOutputBuffer, 0)
    
    ' Write it out. It's possible nothing was actually transformed. However, we will just
    ' rely on the Stream object to optimize for a count of zero bytes to write.
    mStream.WriteBlock mOutputBuffer, 0, TotalBytesTransformed
    
    mBytesBuffered = 0
    mInputIndex = 0
End Sub

''
' This will fill the internal mInputBuffer with as much as possible
' and flush (Transform) it if it was filled.
'
Private Sub WriteToBuffer(ByRef Buffer() As Byte, ByRef Offset As Long, ByRef Count As Long)
    ' Determine how many bytes are needed to fill the buffer. The buffer should
    ' never be full already, so the best we can do is fill it.
    Dim AmountToCopy As Long
    AmountToCopy = CorMath.Min(mInputBlockSize - mBytesBuffered, Count)
    
    ' Don't need to waste our time if there is nothing to do.
    If AmountToCopy = 0 Then
        Exit Sub
    End If
    
    ' Fill our buffer with as much as possible. We still may not have filled it.
    CopyMemory mInputBuffer(mInputIndex), Buffer(Offset), AmountToCopy
    
    ' And we update everyone according to how much we moved.
    mBytesBuffered = mBytesBuffered + AmountToCopy
    mInputIndex = mInputIndex + AmountToCopy
    Offset = Offset + AmountToCopy
    Count = Count - AmountToCopy
    
    ' And if we did fill the buffer, get it transformed and written to the Stream.
    If mBytesBuffered = mInputBlockSize Then
        FlushBuffer
    End If
End Sub

''
' This will write as large a single block of bytes as possible
'
' @param InputBuffer The bytes to be encrypted and written to the Stream.
' @param InputOffset The starting position in InputBuffer to begin encrypting. This is passed ByRef
' to reflect any updates of the position to the calling code.
' @param InputCount The number of bytes to be processed. This is passed ByRef to reflect any updates of the
' count to the calling code.
'
Private Sub WriteMultiBlock(ByRef InputBuffer() As Byte, ByRef InputOffset As Long, ByRef InputCount As Long)
    ' If we have bytes buffered then we need to fill that buffer and flush
    ' it before we do our multiblock transformation.
    If mBytesBuffered > 0 Then
        WriteToBuffer InputBuffer, InputOffset, InputCount
    End If
    
    ' If there isn't even enough bytes to fill a block, then how can we do multiblock?
    ' InputCount can be modified in the WriteToBuffer call.
    If InputCount <= mInputBlockSize Then
        Exit Sub
    End If
    
    ' How many blocks will we be doing at once?
    Dim TotalBytes As Long
    TotalBytes = (InputCount \ mOutputBlockSize) * mInputBlockSize
        
    Dim OutputBuffer() As Byte
    ReDim OutputBuffer(0 To TotalBytes - 1)
    
    ' Apply the tranform to the big array of bytes.
    Dim TotalBytesTransformed As Long
    TotalBytesTransformed = mTransform.TransformBlock(InputBuffer, InputOffset, TotalBytes, OutputBuffer, 0)
    
    ' If anything was transformed, then write it out to the Stream.
    If TotalBytesTransformed > 0 Then
        mStream.WriteBlock OutputBuffer, 0, TotalBytesTransformed
        InputOffset = InputOffset + TotalBytes
        InputCount = InputCount - TotalBytes
    End If
End Sub

Private Function ReadBufferedBytes(ByRef Buffer() As Byte, ByRef Offset As Long, ByRef Count As Long) As Long
    ' If nothing is buffered, we can't read from it
    ' and return what it contains.
    If mBytesBuffered = 0 Then
        Exit Function
    End If
    
    Dim AmountToCopy As Long
    
    ' Only copy what we want or what we already have, which
    ' ever is the smaller value.
    AmountToCopy = CorMath.Min(mBytesBuffered, Count)
    
    ' Make sure there is room in the return buffer.
    If (Offset + AmountToCopy - 1) > UBound(Buffer) Then _
        Error.Argument Argument_SmallConversionBuffer, "Buffer"
    
    ' Copy the data to the return buffer and update all the counters.
    CopyMemory Buffer(Offset), mOutputBuffer(mOutputIndex), AmountToCopy
    Offset = Offset + AmountToCopy
    Count = Count - AmountToCopy
    mBytesBuffered = mBytesBuffered - AmountToCopy
    mOutputIndex = mOutputIndex + AmountToCopy
    ReadBufferedBytes = AmountToCopy
End Function

Private Function ReadMultiBlock(ByRef Buffer() As Byte, ByRef Offset As Long, ByRef Count As Long) As Long
    ' If less than atleast 1 full block is requested, then
    ' we can't perform a mutli-block read, so exit.
    If Count < mOutputBlockSize Then
        Exit Function
    End If
    
    Dim Result As Long
    
    ' Calculate how many full output blocks we want, then determine
    ' how many input blocks in bytes it will take to get it.
    Dim TotalBytes As Long
    TotalBytes = (Count \ mOutputBlockSize) * mInputBlockSize
    
    Dim InputBuffer() As Byte
    ReDim InputBuffer(0 To TotalBytes - 1)
    
    Dim BytesRead As Long
    BytesRead = mStream.ReadBlock(InputBuffer, 0, TotalBytes)
    
    If BytesRead = 0 Then
        Exit Function
    End If
    
    Dim TotalBytesTransformed As Long
    TotalBytesTransformed = mTransform.TransformBlock(InputBuffer, 0, BytesRead, Buffer, Offset)
    
    Offset = Offset + TotalBytesTransformed
    Count = Count - TotalBytesTransformed
    Result = Result + TotalBytesTransformed
    
    ReadMultiBlock = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Stream
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Stream_BeginRead = BeginRead(Buffer, Offset, Count, Callback, State)
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    BeginWrite Buffer, Offset, Count, Callback, State
End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = CanRead
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = CanSeek
End Property

Private Property Get Stream_CanTimeout() As Boolean
    Stream_CanTimeout = CanTimeout
End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = CanWrite
End Property

Private Sub Stream_CloseStream()
    CloseStream
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Stream_EndRead = EndRead(AsyncResult)
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    EndWrite AsyncResult
End Sub

Private Function Stream_Equals(Value As Variant) As Boolean
    Stream_Equals = Equals(Value)
End Function

Private Sub Stream_Flush()
    Flush
End Sub

Private Function Stream_GetHashCode() As Long
    Stream_GetHashCode = GetHashCode
End Function

Private Property Get Stream_Length() As Currency
    Stream_Length = Length
End Property

Private Property Let Stream_Position(ByVal Value As Currency)
    Position = Value
End Property

Private Property Get Stream_Position() As Currency
    Stream_Position = Position
End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    Stream_ReadBlock = ReadBlock(Buffer, Offset, Count)
End Function

Private Function Stream_ReadByte() As Long
    Stream_ReadByte = ReadByte
End Function

Private Property Let Stream_ReadTimeout(ByVal Value As Long)
    ReadTimeout = Value
End Property

Private Property Get Stream_ReadTimeout() As Long
    Stream_ReadTimeout = ReadTimeout
End Property

Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    Stream_SeekPosition = SeekPosition(Offset, Origin)
End Function

Private Sub Stream_SetLength(ByVal Value As Currency)
    SetLength Value
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = ToString
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    WriteBlock Buffer, Offset, Count
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)
    WriteByte Value
End Sub

Private Property Let Stream_WriteTimeout(ByVal Value As Long)
    WriteTimeout = Value
End Property

Private Property Get Stream_WriteTimeout() As Long
    Stream_WriteTimeout = WriteTimeout
End Property
