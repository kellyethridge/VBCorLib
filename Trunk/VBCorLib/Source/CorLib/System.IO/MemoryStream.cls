VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: MemoryStream
'

''
' Provides an in-memory stream of data that can be read/written and modified
' using Stream methods.
'
' @remarks <p>The MemoryStream uses an array of bytes to back the data
' in memory. When the object is first created, no array is allocated until the
' first time the data must be accessed. The default capacity of the array is 4096 bytes.
' The capacity will grow by doubling the size of the current capacity.
' If The capacity needs to be larger, consider setting the Capacity directly to prevent
' the array from having to be resized as new data is written to the stream.</p>
' <p>If the MemoryStream is created using a supplied Byte array, then no internal array
' will be allocated. The MemoryStream will access the supplied Byte array as if it
' were created by the MemoryStream object. As the underlying array is modified, the
' changes can be seen by directly accessing the array. Also, by modifying the array
' directly, the underlying stream will reflect the new data.</p>
' <p><b>CAUTION:</b> If a Byte array is being supplied to the MemoryStream, then the
' array must outlive the MemoryStream object. It is the responsibility of the user to
' destroy the array after the stream has been destroyed. The MemoryStream object
' maintains a reference to the same array in memory, but does not deallocate the
' array when finished. While the object is alive, the array is locked to prevent
' it from being ReDimed or Erased. Closing the stream will release the lock and
' reference to the array, at which point the user has full control of the array.</p>
' <p>The Currency datatypes used are there for consistency across all stream objects.
' Internally, the MemoryStream only supports 2 ^ 31 bytes. Other streams may support more.</p>
'
' @see Constructors
' @see Stream
'
Option Explicit
Implements IObject
Implements Stream

Public Enum BufferMode
    ShareMode
    CopyMode
    AcquireMode
End Enum

Private Const DefaultCapacity As Long = 256

Private mMode           As BufferMode
Private mBufferSet      As Boolean
Private mBuffer()       As Byte
Private mCanWrite       As Boolean
Private mPosition       As Long
Private mCapacity       As Long
Private mLength         As Long
Private mIsClosed       As Boolean
Private mExpandable     As Boolean
Private mIndex          As Long
Private mAsyncPtr       As Long


''
' Returns if the MemoryStream object supports timeout events.
'
' @return Always returns False.
'
Public Property Get CanTimeout() As Boolean
    CanTimeout = False
End Property

''
' Gets the timeout duration for a read operation.
'
' @return The timeout duration in milliseconds.
' @remarks This always throws an InvalidOperationException.
'
Public Property Get ReadTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Sets the timeout duration for a read operation.
'
' @param Value The new duration in milliseconds
' @remarks This always throws an InvalidOperationException.
'
Public Property Let ReadTimeout(ByVal Value As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Gets the timeout duration of a write operation.
'
' @return The timeout duration in milliseconds.
' @remarks This always throws an InvalidOperationException.
'
Public Property Get WriteTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Sets the timeout duration for a write operation.
'
' @param Value The new timeout duration in milliseconds.
' @remarks This always throws an InvalidOperationException.
'
Public Property Let WriteTimeout(ByVal Value As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Begins an asynchronous buffer read. For a MemoryStream, calling an asynchronous buffer
' read is identical to calling ReadBlock.
'
' @param Buffer The destination array to copy the bytes into.
' @param Offset The index in Buffer to begin copying to.
' @param Count The maximum number of bytes to copy.
' @param Callback A listening object to be notified that the read is finished.
' @param State User data that can be carried within the IAsyncResult object returned.
' @return An IAsyncResult object used to identify the corrisponding Endread and report
' information about the buffer read.
' @remarks <p>Since the stream is directly stored and accessed in memory, there is no
' waiting for reads to occur. The BeginRead method works identical to the Read method
' in terms of immediately reading stream data into the <i>Buffer</i>.</p>
' <p>When defining the Count for the number of maximum bytes to be read, if the
' stream does not have enough bytes, the read will copy less bytes than Count.</p>
' <p>If an exception has occurred during the read, it will not be thrown until the
' EndRead is executed.
'
Public Function BeginRead(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Result As StreamAsyncResult
    Set Result = Cor.NewStreamAsyncResult(State)
    
    On Error GoTo Catch
    Result.BytesRead = ReadBlock(Buffer, Offset, Count)
    Result.IsCompleted = True
    Result.IsReadType = True
    
    If Not Callback Is Nothing Then
        Callback.Execute Result
    End If

Catch:
    Dim Ex As Exception
    If Catch(Ex, Err) Then
        Set Result.Exception = Ex
    End If
    
    mAsyncPtr = ObjPtr(Result)
    Set BeginRead = Result
End Function

''
' Begins an asynchronous buffer write. For a MemoryStream, having an asynchronous buffer
' write is of no use.
'
' @param Buffer The source array to copy the bytes from into the stream.
' @param Offset The starting index in the source <i>Buffer</i> to begin copying from.
' @param Count The maximum number of bytes to be copied into the stream.
' @param Callback A user supplied object to be notified of the finish of the writing.
' @param State User data that can be carried within the IAsyncResult object return.
' @return An IAsyncResult object used to identify the corrisponding EndBread and report
' information about the buffer read.
' @remarks <p>Since the stream is directly stored and accessed in memory, there is no
' waiting for reads to occur. The BeginWrite method works identical to the WriteBlock method
' in terms of reading immediately writing stream data to the stream.</p>
' <p>When defining the Count for the number of maximum bytes to be written, if the
' <i>Buffer</i> does not have enough bytes, an exception will be thrown in the EndWrite method.</p>
' <p>If an exception has occurred during the read, it will not be thrown until the
' EndRead is executed.
'
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Result As StreamAsyncResult
    Set Result = Cor.NewStreamAsyncResult(State)
    
    On Error GoTo Catch
    Result.CompletedSynchronously = True
    Result.IsCompleted = True
    WriteBlock Buffer, Offset, Count
    
    If Not Callback Is Nothing Then
        Callback.Execute Result
    End If
    
Catch:
    Dim Ex As Exception
    If Catch(Ex, Err) Then
        Set Result.Exception = Ex
    End If
    
    mAsyncPtr = ObjPtr(Result)
    Set BeginWrite = Result
End Function

''
' Signifies the end of an asynchronous read from the stream.
'
' @param AsyncResult The IAsyncResult object returned from the BeginRead function.
' @return The number of bytes read from the stream in to <i>Buffer</i> in the BeginRead method.
' @remarks <p>This method is generally called anytime after BeginRead for a MemoryStream.
' Since the MemoryStream uses memory for the data store, there is no need to wait for
' the data to arrive. The BeginRead/EndRead calls are redundant and should not be
' used with the MemoryStream. Instead use the Read Method.</p>
' <p>If any exceptions were caused during the reading then they will be thrown
' during the call to EndRead.</p>
'
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long
    If AsyncResult Is Nothing Then _
        Error.ArgumentNull "AsyncResult"
    
    Dim Result As StreamAsyncResult
    Set Result = StrongPtr(mAsyncPtr)
    If Result Is Nothing Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple, "AsyncResult"
    If Not Result Is AsyncResult Then _
        Error.InvalidOperation InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple
    If Not Result.IsReadType Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple, "AsyncResult"
    
    mAsyncPtr = vbNullPtr
    If Not Result.Exception Is Nothing Then
        Throw Result.Exception
    End If
    
    EndRead = Result.BytesRead
End Function

''
' Signifies the end of an asynchronous write to the stream.
'
' @param AsyncResult The IAsyncResult returned from the BeginWrite function.
' @remarks <p>This method is generally called anytime after BeginWrite for a MemoryStream.
' Since the MemoryStream uses memory for the data store, there is no need to wait for
' the data to be written. The BeginWrite/EndWrite calls are redundant and should not be
' used with the MemoryStream. Instead use the WriteBlock Method.</p>
' <p>If any exceptions were caused during the reading then they will be thrown
' during the call to EndRead.</p>
'
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult)
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult")
    
    Dim Result As StreamAsyncResult
    Set Result = StrongPtr(mAsyncPtr)
    If Result Is Nothing Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple, "AsyncResult"
    If Not Result Is AsyncResult Then _
        Error.InvalidOperation InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple
    If Result.IsReadType Then _
        Error.Argument InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple, "AsyncResult"
    
    mAsyncPtr = vbNullPtr
    If Not Result.Exception Is Nothing Then
        Throw Result.Exception
    End If
End Sub

''
' Determines if the current MemoryStream can be read.
'
' @return The status of the ability to read the stream.
' @remarks If the MemoryStream is open, then it can be read. If it is closed, it cannot be read.
'
Public Property Get CanRead() As Boolean
    CanRead = Not mIsClosed
End Property

''
' Determines if the current MemoryStream can be written to.
'
' @return The status of the ability to write to the stream.
' @remarks If the stream was created from a user-supplied array, then the user
' will determine if the stream is allowed to be written to. Otherwise, the stream
' is always writable while it is open.
'
Public Property Get CanWrite() As Boolean
    CanWrite = mCanWrite And (Not mIsClosed)
End Property

''
' Determines if the current MemoryStream can seek positions within the stream.
'
' @return The status of the ability to seek a position within the stream.
' @remarks If the MemoryStream is open, then it is seekable, otherwise it is not.
'
Public Property Get CanSeek() As Boolean
    CanSeek = Not mIsClosed
End Property

''
' Returns the current capacity of the internal stream buffer.
'
' @return The number of bytes the internal array contains.
' @remarks If the the capacity is reached while writing, then the capacity
' will be expanded, unless the underlying byte array is user-supplied. In
' this case, an exception will be thrown if the capacity is changed.
'
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the current capacity for the internal stream buffer.
'
' @param Value The number of bytes to allocate for the internal array.
' @remarks The capacity cannot be set less than the current length of
' the stream. If the underlying byte array is user-supplied, then an
' exception will be thrown if the capacity is modified.
'
Public Property Let Capacity(ByVal Value As Long)
    If Not mExpandable Then _
        Error.NotSupported NotSupported_MemoryStreamNotExpandable
    If Value < mLength Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_SmallCapacity
    
    EnsureOpen
    
    If Value = mLength Then
        Exit Property
    End If
    
    mCapacity = Value
    
    If mBufferSet Then
        ReDim Preserve mBuffer(0 To mCapacity - 1)
    Else
        EnsureBuffer
    End If
End Property

''
' Returns the current length of the stream.
'
' @return The number of bytes the current stream is.
' @remarks The length of the stream differs from the capacity in that the
' capacity is the amount of allocated memory used to store the stream, and the
' length is the number of bytes used from within the allocated memory.
'
Public Property Get Length() As Currency
    Length = mLength
End Property

''
' Returns the current position within the stream.
'
' @return The byte index within the stream.
' @remarks The position is the index at which the next read or write
' will occur. Position is alway zero based, even if a user-supplied
' array is not.
'
Public Property Get Position() As Currency
    Position = mPosition
End Property

''
' Sets the current position within the stream.
'
' @param Value The absolution position within the stream.
' @remarks The position is the index at which the next read or write
' will occur. Position is always zero based, even if a user-supplied
' array is not.
'
Public Property Let Position(ByVal Value As Currency)
    Value = Int(Value)
    If Value < 0 Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_NeedNonNegNum
    
    EnsureOpen
    mPosition = Value
End Property

''
' Closes the MemoryStream and releases any resources allocated.
'
' @remarks If a user-supplied array is being used, it will be released
' at this point. The user then may modify the array as normal.
'
Public Sub CloseStream()
    If mIsClosed Then Exit Sub
    
    If Not mExpandable And (mMode = ShareMode) Then
        SafeArrayUnlock SAPtr(mBuffer)
        SAPtr(mBuffer) = vbNullPtr
    End If
    
    mIsClosed = True
End Sub

''
' Writes an array of bytes to the underlying stream.
'
' @param Buffer The bytes to be written to the stream.
' @param Offset The startind index in <i>Buffer</i> to begin writing from.
' @param Count The number of bytes to be written to the stream.
' @remarks <p>If the position in the stream is closer to the end than Count bytes, then
' the stream will expand if possible. If the underlying stream is a user-supplied
' byte array, then an exception will occur if the capacity is reached.</p>
' <p>The stream must be open to write to it, otherwise an exception will be thrown.</p>
'
Public Sub WriteBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    ValidateByteArrayRange Buffer, Offset, Count, Parameter_Buffer, Parameter_Offset
    EnsureWritable
    EnsureOpen
    EnsureBuffer
    
    If Count = 0 Then
        Exit Sub
    End If
    
    EnsureCapacity mPosition + Count
    CopyMemory mBuffer(mIndex + mPosition), Buffer(Offset), Count
    mPosition = mPosition + Count
    
    If mPosition > mLength Then
        mLength = mPosition
    End If
End Sub

''
' Writes a single byte to the underlying stream.
'
' @param value The byte to be written to the stream.
' @remarks <p>If the position to be written exceeds the current capacity, then
' the stream will be expanded if possible. If the underlying stream is a
' user-supply byte array, then the capacity cannot be modified and an
' exception will be thrown.</p>
' <p>The stream must be open to write to it, otherwise an exception will be thrown.</p>
'
Public Sub WriteByte(ByVal Value As Byte)
    EnsureWritable
    EnsureOpen
    EnsureBuffer
    
    mBuffer(mPosition + mIndex) = Value
    mPosition = mPosition + 1
    
    If mPosition > mLength Then
        mLength = mPosition
    End If
End Sub

''
' Reads a specifiec number of bytes from the stream.
'
' @param Buffer The destination to store the bytes being read from the stream.
' @param Offset The startind index in <i>Buffer</i> to begin storing the bytes read from the stream.
' @param Count The maximum number of bytes to be read from the stream.
' @return The total number of bytes read from the stream.
' @remarks <p>If Offset plus Count exceeds the end of <i>Buffer</i> then an exception will be thrown.
' If the current position in the stream is closer to the end than Count bytes, then only the
' remaining bytes will be read.</p>
' <p>The stream must be open to read from it, otherwise an exception will be thrown.</p>
' <p>After reading from the stream, the current position will advance the number of byte as returend.</p>
'
Public Function ReadBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    ValidateByteArrayRange Buffer, Offset, Count, Parameter_Buffer, Parameter_Offset
    EnsureOpen
    EnsureBuffer
    
    If mPosition >= mLength Then Exit Function
    If mPosition + Count > mLength Then
        Count = mLength - mPosition
    End If
    
    CopyMemory Buffer(LBound(Buffer) + Offset), mBuffer(mPosition + mIndex), Count
    mPosition = mPosition + Count
    ReadBlock = Count
End Function

''
' Reads a single byte from the stream.
'
' @return The byte value from the current position in the stream.
' @remarks <p>The stream must be open for read from it, otherwise an exception will be thrown.</p>
' <p>If the end of the stream is reached during a read, then a value of
' -1 is returned.</p>
' <p>Each read from the stream will advance the position one byte.</p>
'
Public Function ReadByte() As Long
    If mPosition >= mLength Then
        ReadByte = -1
        Exit Function
    End If
    
    ReadByte = mBuffer(mPosition + mIndex)
    mPosition = mPosition + 1
End Function

''
' Changes the position within the stream.
'
' @param Offset The number of bytes to move.
' @param Loc The starting position in the stream to move from.
' @return The actual position moved to.
' @param <p>The offset is applied to a starting position refered to by <i>Loc</i>.
' the <i>Offset</i> can be negative to move the position backwards.</p>
' <p>If the position ends up being before the start of the stream an exception will be thrown.
' The position can be set passed the end of the stream. This will cause the stream to be
' expanded on the next access. If the stream expansion is not supported an exception
' will be thrown. This may be the case when using a user-supplied byte array.</p>
'
Public Function SeekPosition(ByVal Offset As Currency, ByVal Loc As SeekOrigin) As Currency
    Offset = Int(Offset)
    Select Case Loc
        Case FromBeginning
            Position = Offset
        Case FromCurrent
            Position = mPosition + Offset
        Case FromEnd
            Position = mLength + Offset
        Case Else
            Error.Argument Argument_InvalidSeekOrigin
    End Select
    SeekPosition = mPosition
End Function

''
' Sets the current length of the stream.
'
' @param value The size of the stream in bytes.
' @remarks <p>The shortening the length of the stream will truncate the excess data. If
' increasing the length, then the new exposed data will be set to zero.</p>
' <p>If a user-supplied byte array is being used, this function has limited capabilites.
' The size of the stream can be changed to withing the origin size of the user-supplied
' byte array. If a length greater than the original array size is set, then an exception
' is thrown.</p>
'
Public Sub SetLength(ByVal Value As Currency)
    EnsureWritable
    
    Value = Int(Value)
    Select Case Value
        Case mLength + 1 To mCapacity
            Dim i As Long
            For i = mLength To Value - 1
                mBuffer(mIndex + i) = 0
            Next i
        Case Is > mCapacity
            Capacity = Value
        Case Is < 0
            Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_NeedNonNegNum
    End Select
    
    mLength = Value
End Sub

''
' Returns a byte array copy of the underlying byte stream.
'
' @return The byte array copy of the underlying stream.
' @remarks If the underlying stream is a user-supplied byte array, then only
' that portion determined by Count when the stream was created is returned, not
' the entire array.
'
Public Function ToArray() As Byte()
    EnsureBuffer
    
    If mLength > 0 Then
        Dim Result() As Byte
        ReDim Result(0 To mLength - 1)
        CopyMemory Result(0), mBuffer(mIndex), mLength
        ToArray = Result
    Else
        ToArray = Cor.NewBytes
    End If
End Function

''
' This has no purpose in a MemoryStream
'
Public Sub Flush()
    ' just here for consistency
End Sub

''
' Writes the underlying stream to another stream.
'
' @param Stream The stream to write the underlying stream to.
'
Public Sub WriteTo(ByVal Stream As Stream)
    If Stream Is Nothing Then _
        Error.ArgumentNull "Stream", ArgumentNull_Stream
    EnsureBuffer
    EnsureOpen
    
    Stream.WriteBlock mBuffer, mIndex, mLength
End Sub


''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mExpandable = True
    mCapacity = DefaultCapacity
    mCanWrite = True
End Sub

Friend Sub Init(ByRef Buffer() As Byte, ByRef Index As Variant, ByRef Count As Variant, ByVal Writable As Boolean, ByVal Mode As BufferMode)
    Dim Range As ListRange
    Range = ValidateByteArrayOptionalRange(Buffer, Index, Count, Parameter_Buffer)
        
    mIndex = Range.Index
    mCapacity = Range.Count
    mCanWrite = Writable
    mMode = Mode
    mLength = mCapacity
    
    Select Case Mode
        Case BufferMode.ShareMode
            SAPtr(mBuffer) = SAPtr(Buffer)
            SafeArrayLock SAPtr(mBuffer)
            mExpandable = False
        Case BufferMode.CopyMode
            mBuffer = Buffer
        Case BufferMode.AcquireMode
            SAPtr(mBuffer) = SAPtr(Buffer)
            SAPtr(Buffer) = vbNullPtr
        Case Else
            Throw Cor.NewArgumentException("Invalid buffer mode.")
    End Select
    
    mBufferSet = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureBuffer()
    If Not mBufferSet Then
        ReDim mBuffer(0 To mCapacity - 1)
        mCanWrite = True
        mBufferSet = True
    End If
    
    If mExpandable Then
        EnsureCapacity mPosition + 1
    End If
End Sub

Private Sub EnsureOpen()
    If mIsClosed Then
        Throw Cor.NewObjectDisposedException("MemoryStream", Environment.GetResourceString(ObjectDisposed_StreamClosed))
    End If
End Sub

Private Sub EnsureWritable()
    If Not mCanWrite Then
        Throw Cor.NewNotSupportedException(Environment.GetResourceString(NotSupported_UnwritableStream))
    End If
End Sub

Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    
    If RequiredCapacity <= mCapacity Then Exit Sub
    NewCapacity = mCapacity * 2
    If NewCapacity < RequiredCapacity Then
        NewCapacity = RequiredCapacity + DefaultCapacity
    End If
    
    Capacity = NewCapacity
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mBuffer = .ReadProperty("Buffer")
        mCanWrite = .ReadProperty("CanWrite")
        mPosition = .ReadProperty("Position")
        mCapacity = .ReadProperty("Capacity")
        mLength = .ReadProperty("Length")
        mIsClosed = .ReadProperty("IsClosed")
        mIndex = .ReadProperty("Index")
    End With
End Sub

Private Sub Class_Terminate()
    CloseStream
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Buffer", mBuffer
        .WriteProperty "CanWrite", mCanWrite
        .WriteProperty "Position", mPosition
        .WriteProperty "Capacity", mCapacity
        .WriteProperty "Length", mLength
        .WriteProperty "IsClosed", mIsClosed
        .WriteProperty "Index", mIndex
    End With
End Sub


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
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Set Stream_BeginRead = BeginRead(Buffer, Offset, Count, Callback, State)
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Set Stream_BeginWrite = BeginWrite(Buffer, Offset, Count, Callback, State)
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
    Call WriteByte(Value)
End Sub

Private Property Let Stream_WriteTimeout(ByVal Value As Long)
    WriteTimeout = Value
End Property

Private Property Get Stream_WriteTimeout() As Long
    Stream_WriteTimeout = WriteTimeout
End Property
