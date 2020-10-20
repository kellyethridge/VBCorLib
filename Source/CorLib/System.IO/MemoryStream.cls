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
' using <a href="Stream.html">Stream</a> methods.
'
' @remarks <p>The <b>MemoryStream</b> uses an array of bytes to back the data
' in memory. When the object is first created, no array is allocated until the
' first time the data must be accessed. The default capacity of the array is 4096 bytes.
' The capacity will grow by doubling the size of the current capacity.
' If The capacity needs to be larger, consider setting the <a href="MemoryStream.Let.Capacity.html">Capacity</a> directly to prevent
' the array from having to be resized as new data is written to the stream.</p>
' <p>If the <b>MemoryStream</b> is created using a supplied <b>Byte</b> array, then no internal array
' will be allocated. The <b>MemoryStream</b> will access the supplied <b>Byte</b> array as if it
' were created by the <b>MemoryStream</b> object. As the underlying array is modified, the
' changes can be seen by directly accessing the array. Also, by modifying the array
' directly, the underlying stream will reflect the new data.</p>
' <p>The <b>Currency</b> type is used for consistency across all stream objects.
' Internally, the <b>MemoryStream</b> only supports 2 ^ 31 bytes. Other streams may support more.</p>
' @caution
' If a <b>Byte</b> array is being supplied to the <b>MemoryStream</b> and <i>Mode</i> is <a href="MemoryStream.BufferMode.html">ShareMode</a>,
' then the array must out-live the <b>MemoryStream</b> object. It is the responsibility of the user to
' destroy the array after the stream has been destroyed. The <b>MemoryStream</b> object
' maintains a reference to the same array in memory, but does not deallocate the
' array when finished. While the object is alive, the array is locked to prevent
' it from being ReDimed or Erased. Closing the stream will release the lock and
' reference to the array, at which point the user has full control of the array.
' @see Constructors
' @see Stream
' @include "..\DocIncludes\MemoryStream.html"
'@Folder("CorLib.System.IO")
Option Explicit
Implements IObject
Implements Stream

''
' The method a specified byte array should be referenced.
'
' @param ShareMode The <a href="MemoryStream.html">MemoryStream</a> points to the original array. This prevents a copy of the
' array from occurring. Thid mode requires the specified array to out-live the <a href="MemoryStream.html">MemoryStream</a> object instance.
' @param CopyMode A new array will be allocated with the contents of the specified array leaving the original array intact.
' @param AcquireMode An internal array will point to the original array and set the original array variable to null. This effectively steals the specified array.
Public Enum BufferMode
    ShareMode
    CopyMode
    AcquireMode
End Enum

Private Const DefaultCapacity As Long = 256
Private Const MaxStreamLength As Long = 2147483647

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
' Gets if the <a href="MemoryStream.html">MemoryStream</a> object supports timeout events.
'
' @return Always returns <b>False</b>.
Public Property Get CanTimeout() As Boolean
    CanTimeout = False
End Property

''
' Gets the timeout duration for a read operation.
'
' @return The timeout duration in milliseconds.
' @remarks This always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.
Public Property Get ReadTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Sets the timeout duration for a read operation.
'
' @param Value The new duration in milliseconds
' @remarks This always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.
Public Property Let ReadTimeout(ByVal Value As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Gets the timeout duration of a write operation.
'
' @return The timeout duration in milliseconds.
' @remarks This always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.
Public Property Get WriteTimeout() As Long
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Sets the timeout duration for a write operation.
'
' @param Value The new timeout duration in milliseconds.
' @remarks This always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.
Public Property Let WriteTimeout(ByVal Value As Long)
    Error.InvalidOperation InvalidOperation_Timeouts
End Property

''
' Begins an asynchronous buffer read. For a <a href="MemoryStream.html">MemoryStream</a>, calling an asynchronous buffer
' read is identical to calling <a href="MemoryStream.ReadBlock.html">ReadBlock</a>.
'
' @param Buffer The destination array to copy the bytes into.
' @param Offset The index in <i>Buffer</i> to begin copying to.
' @param Count The maximum number of bytes to copy.
' @param Callback A listening object to be notified that the read is finished.
' @param State User data that can be carried within the <a href="IAsyncResult.html">IAsyncResult</a> object returned.
' @return An <a href="IAyncResult.html">IAsyncResult</a> object used to identify the corrisponding <a href="MemoryStream.EndRead.html">Endread</a> and report
' information about the buffer read.
' @remarks <p>Since the stream is directly stored and accessed in memory, there is no
' waiting for reads to occur. The <b>BeginRead</b> method works identical to the <a href="MemoryStream.ReadBlock.html">ReadBlock</a> method
' in terms of immediately reading stream data into the <i>Buffer</i>.</p>
' <p>When defining the Count for the number of maximum bytes to be read, if the
' stream does not have enough bytes, the read will copy less bytes than Count.</p>
' <p>If an exception has occurred during the read, it will not be thrown until the
' EndRead is executed.</p>
Public Function BeginRead(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Result As StreamAsyncResult
    
    If Not CanRead Then _
        Error.NotSupported Argument_StreamNotReadable
    
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
    
    mAsyncPtr = WeakPtr(Result)
    Set BeginRead = Result
End Function

''
' Begins an asynchronous buffer write. For a <a href="MemoryStream.html">MemoryStream</a>, having an asynchronous buffer write is of no use.
'
' @param Buffer The source array to copy the bytes from into the stream.
' @param Offset The starting index in the source <i>Buffer</i> to begin copying from.
' @param Count The maximum number of bytes to be copied into the stream.
' @param Callback A user supplied object to be notified of the finish of the writing.
' @param State User data that can be carried within the <a href="IAsyncResult.html">IAsyncResult</a> object return.
' @return An <a href="IAsyncResult.html">IAsyncResult</a> object used to identify the corrisponding <a href="MemoryStream.EndRead.html">EndBread<a/> and report
' information about the buffer read.
' @remarks <p>Since the stream is directly stored and accessed in memory, there is no
' waiting for reads to occur. The <b>BeginWrite</b> method works identical to the <a href="MemoryStream.WriteBlock.html">WriteBlock</a> method
' in terms of reading immediately writing stream data to the stream.</p>
' <p>When defining the <i>Count</i> for the number of maximum bytes to be written, if the
' <i>Buffer</i> does not have enough bytes, an exception will be thrown in the <a href="MemoryStream.EndWrite.html">EndWrite</a> method.</p>
' <p>If an exception has occurred during the read, it will not be thrown until the <a href="MemoryStream.EndRead.html">EndRead</a> is executed.</p>
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByRef State As Variant) As IAsyncResult
    Dim Result As StreamAsyncResult
    
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    
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
    
    mAsyncPtr = WeakPtr(Result)
    Set BeginWrite = Result
End Function

''
' Signifies the end of an asynchronous read from the stream.
'
' @param AsyncResult The <a href="IAsyncResult.html">IAsyncResult</a> object returned from the <a href="MemoryStream.BeginRead.html">BeginRead</a> function.
' @return The number of bytes read from the stream in to <i>Buffer</i> in the BeginRead method.
' @remarks <p>This method is generally called anytime after <a href="MemoryStream.BeginRead.html">BeginRead</a> for a <a href="MemoryStream.html">MemoryStream</a>.
' Since the <a href="MemoryStream.html">MemoryStream</a> uses memory for the data store, there is no need to wait for
' the data to arrive. The <a href="MemoryStream.BeginRead.html">BeginRead</a>/<a href="MemoryStream.EndRead.html">EndRead</a> calls are redundant and should not be
' used with the <a href="MemoryStream.html">MemoryStream</a>. Instead use the <a href="MemoryStream.ReadBlock.html">ReadBlock</a> Method.</p>
' <p>If any exceptions were caused during the reading then they will be thrown during the call to <a href="MemoryStream.EndRead.html">EndRead</a>.</p>
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
' @param AsyncResult The <a href="IAsyncResult.html">IAsyncResult<a/> returned from the <a href="MemoryStream.BeginWrite.html">BeginWrite</a> function.
' @remarks <p>This method is generally called anytime after <a href="MemoryStream.BeginWrite.html">BeginWrite</a> for a <a href="MemoryStream.html">MemoryStream</a>.
' Since the <a href="MemoryStream.html">MemoryStream</a> uses memory for the data store, there is no need to wait for
' the data to be written. The <a href="MemoryStream.BeginWrite.html">BeginWrite</a>/<a href="MemoryStream.EndWrite.html">EndWrite</a> calls are redundant and should not be
' used with the <a href="MemoryStream.html">MemoryStream</a>. Instead use the <a href="MemoryStream.WriteBlock.html">WriteBlock</a> Method.</p>
' <p>If any exceptions were caused during the reading then they will be thrown during the call to <a href="MemoryStream.EndRead.html">EndRead</a>.</p>
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
' Gets value indicating whether the current stream supports reading.
'
' @return <b>True</b> if the stream supports reading; otherwise, <b>False</b>.
' @remarks If the stream is closed, this property return <b>False</b>.
Public Property Get CanRead() As Boolean
    CanRead = Not mIsClosed
End Property

''
' Gets value indicating whether the current stream supports writing.
'
' @return <b>True</b> if the stream supports writing; otherwise, <b>False</b>.
' @remarks If the <a href="MemoryStream.html">MemoryStream</a> object was created with the <i>Writable</i>
' parameter specified as <b>True</b> or the object was created with no parameters, then the stream supports writing.
' <p>If the stream is closed, this property return <b>False</b>.</p>
Public Property Get CanWrite() As Boolean
    CanWrite = mCanWrite And (Not mIsClosed)
End Property

''
' Gets value indicating whether the current stream supports seeking.
'
' @return <b>True</b> if the stream supports seeking; otherwise, <b>False</b>.
' @remarks If the stream is closed, this property return <b>False</b>.
Public Property Get CanSeek() As Boolean
    CanSeek = Not mIsClosed
End Property

''
' Gets the number of bytes allocated for this stream.
'
' @return The length of the usable portion of the buffer for the stream.
' @include "..\DocIncludes\MemoryStream.Capacity.html"
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the number of bytes allocated for this stream.
'
' @param Value The length of the usable portion of the buffer for the stream.
' @include "..\DocIncludes\MemoryStream.Capacity.html"
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
' Gets the length of the stream in bytes.
'
' @return The length of the stream in bytes.
' @remarks The length of the stream differs from the capacity in that the capacity is the amount of allocated
' memory used to store the stream, and the length is the number of bytes used from within the allocated memory.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'    </tr>
' </table>
Public Property Get Length() As Currency
    EnsureOpen
    Length = mLength
End Property

''
' Returns the current position within the stream.
'
' @return The byte index within the stream.
' @remarks The position is the index at which the next read or write will occur. Position is alway zero based, even if a user-supplied array is not.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'    </tr>
' </table>
Public Property Get Position() As Currency
    EnsureOpen
    Position = mPosition
End Property

''
' Sets the current position within the stream.
'
' @param Value The current position within the stream.
' @remarks The position is the index at which the next read or write will occur. Position is always zero based, even if a user-supplied array is not.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>The position is set to a negative value or a value greater than 32bit integer max value.</td>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'    </tr>
' </table>
Public Property Let Position(ByVal Value As Currency)
    Value = Int(Value)
    If Value < 0 Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_NeedNonNegNum
    
    EnsureOpen
    
    If Value > MaxStreamLength Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_StreamLength
        
    mPosition = Value
End Property

''
' Closes the MemoryStream and releases any resources allocated.
'
' @remarks If a user-supplied array is being used, it will be released
' at this point. The user then may modify the array as normal.
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
' @param Offset The starting index in <i>Buffer</i> to begin writing from.
' @param Count The maximum number of bytes to be written to the stream.
' @remarks The <i>Offset</i> parameter gives the offset of the first byte in <i>Buffer</i> to write from, and the <i>Count</i> parameter gives
' the number of bytes to write. If the write operation is successful, the current position within the stream is advanced
' by the number of bytes written. If an exception occurs, the current position within the stream is unchanged.
' <p>Except for a <a href="MemoryStream.html">MemoryStream</a> constructed with a Byte() parameter, write operations at the end of a <a href="MemoryStream.html">MemoryStream</a> expand the <a href="MemoryStream.html">MemoryStream</a>.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Buffer</i> is uninitialized.</td>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The stream does not support writing. For additional information see <a href="MemoryStream.Get.CanWrite.html">CanWriter</a>.<br/>
'           -or-<br/>
'           The current position is closer than <i>Count</i> bytes to the end of the stream, and the capacity cannot be modified.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Offset</i> subtracted from the <i>Buffer</i> length is less than <i>Count</i>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Offset</i> or <i>Count</i> is negative.</td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The current stream instance is closed.</td>
'    </tr>
' </table>
Public Sub WriteBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    ValidateArrayRange Buffer, Offset, Count, NameOfBuffer, NameOfOffset
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
' @remarks Except for a <a href="MemoryStream.html">MemoryStream</a> constructed with a Byte() parameter,
' write operations at the end of a <a href="MemoryStream.html">MemoryStream</a> expand the <a href="MemoryStream.html">MemoryStream</a>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The stream does not support writing. For additional information see <a href="MemoryStream.Get.CanWrite.html">CanWriter</a>.<br/>
'           -or-<br/>
'           The current position is at the end of the stream, and the capacity cannot be modified.
'        </td>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The current stream instance is closed.</td>
'    </tr>
' </table>
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
' If the current position in the stream is closer to the end than <i>Count</i> bytes, then only the remaining bytes will be read.</p>
' <p>The stream must be open to read from it, otherwise an exception will be thrown.</p>
' <p>After reading from the stream, the current position will advance the number of byte as returend.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Buffer</i> is uninitialized.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Offset</i> subtracted from the <i>Buffer</i> length is less than <i>Count</i>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Offset</i> or <i>Count</i> is negative.</td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The current stream instance is closed.</td>
'    </tr>
' </table>
Public Function ReadBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    ValidateArrayRange Buffer, Offset, Count, NameOfBuffer, NameOfOffset
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
' @return The byte cast to a <b>Long</b>, or -1 if the end of the stream has been reached.
' @remarks If the read operation is successful, the current position within the stream is advanced by one byte.
' If an exception occurs, the current position within the stream is unchanged.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The current stream instance is closed.</td>
'    </tr>
' </table>
Public Function ReadByte() As Long
    EnsureOpen
    
    If mPosition >= mLength Then
        ReadByte = -1
        Exit Function
    End If
    
    ReadByte = mBuffer(mPosition + mIndex)
    mPosition = mPosition + 1
End Function

''
' Sets the position within the current stream to the specified value.
'
' @param Offset The new position within the stream. This is relative to the <i>Loc</i> parameter, and can be positive or negative.
' @param Loc A value of type <a href="Stream.SeekOrigin.html">SeekOrigin</a>, which acts as the seek reference point.
' @return The new position within the stream, calculated by combining the initial reference point and the offset.
' @remarks Seeking to any location beyond the length of the stream is supported.
' <p>Do not use the <b>SeekPosition</b> method to determine the new position in the stream if the <a href="MemoryStream.html">MemoryStream</a> was initialized with
' a non-zero offset. If you do, <b>SeekPosition</b> will return an inaccurate value. Instead, use the <a href="MemoryStream.Get.Position.html">Position</a> property to get the new position of the stream.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Offset</i> is greater than max 32-bit value.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td>
'        <td>There is an invalid <a href="Stream.SeekOrigin.html">SeekOrigin</a>.</td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The current stream instance is closed.</td>
'    </tr>
' </table>
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The current stream is not resizable and <i>Value</i> is larger than the current capacity.<br/>
'           -or-<br/>
'           The current stream does not support writing.
'        </td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>value is negative or is greater than the maximum length of the MemoryStream, where the maximum
'            length is( MaxValue - origin), and origin is the index into the underlying buffer at which the stream starts.</td>
'    </tr>
' </table>
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
' Writes the stream contents to a byte array, regardless of the <a href="MemoryStream.Get.Position.html">Position</a> property.
'
' @return A new byte array.
' @remarks This method omits unused bytes in <a href="MemoryStream.html">MemoryStream</a> from the array.
' <p>This method returns a copy of the contents of the <a href="MemoryStream.html">MemoryStream</a> as a byte array.
' If the current instance was constructed on a provided byte array, a copy of the section of the array to which
' this instance has access is returned. See the <a href="MemoryStream.html">MemoryStream</a> constructor <a href="Constructors.NewMemoryStream.html">NewMemoryStream</a> for details.
' <table style="width: auto">
'   <tr>
'       <th>Note</td>
'   </tr>
'   <tr>
'       <td>The method works when the <a href="MemoryStream.html">MemoryStream</a> is closed.</td>
'   </tr>
' </table>
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
Public Sub Flush()
    ' just here for consistency
End Sub

''
' Writes the entire contents of this memory stream to another stream.
'
' @param Stream The stream to write this memory stream to.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Stream</i> is <b>Nothing</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The current or target stream is closed.</td>
'    </tr>
' </table>
Public Sub WriteTo(ByVal Stream As Stream)
    If Stream Is Nothing Then _
        Error.ArgumentNull "Stream", ArgumentNull_Stream
    EnsureBuffer
    EnsureOpen
    
    Stream.WriteBlock mBuffer, mIndex, mLength
End Sub

''
' Reads the bytes from the current stream and writes them to another stream.
'
' @param Destination The stream to which the contents of the current stream will be copied.
' @param BufferSize The size of the buffer. This value must be greater than zero. The default size is 81920.
' @remarks Copying begins at the current position in the current stream, and does not reset the position of the destination stream after the copy operation is complete.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Destination</i> is <b>Nothing</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The current stream does not support reading.<br/>
'           -or-<br/>
'           <i>Destination</i> does not support writing.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>Either the current stream or <i>Destination</i> were closed before the CopyTo method was called.</td>
'    </tr>
'    <tr>
'        <td><a href="IOException.html">IOException</a></td><td>An I/O error occurred.</td>
'    </tr>
' </table>
Public Sub CopyTo(ByVal Destination As Stream, Optional ByVal BufferSize As Long = 81920)
    Statics.Stream.CopyTo Me, Destination, BufferSize
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
    ValidateArrayRange Buffer, Index, Count, NameOfBuffer
    
    Dim Range As ListRange

    Range = MakeArrayRange(Buffer, Index, Count)
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
            Error.ArgumentOutOfRange ArgumentOutOfRange_Enum, "Mode"
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

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
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

Private Sub Stream_CopyTo(ByVal Destination As Stream, Optional ByVal BufferSize As Long = 81920)
    CopyTo Destination, BufferSize
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
