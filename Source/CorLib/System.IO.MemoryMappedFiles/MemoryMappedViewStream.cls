VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedViewStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: MemoryMappedViewStream
'

''
' Represents a view of a memory-mapped file as a sequentially accessed stream.
Option Explicit
Implements IObject
Implements Stream

Private mMappingHandle  As SafeMemoryMappedFileHandle
Private mViewAddress    As Long
Private mCapacity       As Long
Private mViewStream     As MemoryStream
Private mViewData()     As Byte
Private mViewDataSA     As SafeArray1d
Private mCanRead        As Boolean
Private mCanWrite       As Boolean


''
' Gets a value indicating whether a stream supports reading.
'
' @return <b>False</b> if the object was created by a constructor with an <i>Access</i> parameter that did not include reading the stream and if the stream is closed; otherwise, <b>True</b>.
' @remarks This property indicates whether the current stream object supports reading.
Public Property Get CanRead() As Boolean
    CanRead = mCanRead And mViewStream.CanRead
End Property

''
' Gets a value indicating whether a stream supports writing.
'
' @return <b>False</b> if the object was created by a constructor with an <b>Access</b> parameter value that supports writing or was created by a constructor that had no parameters, or if the stream is closed; otherwise, <b>True</b>.
' @remarks This property indicates whether the current stream object supports writing.
Public Property Get CanWrite() As Boolean
    CanWrite = mCanWrite And mViewStream.CanWrite
End Property

''
' Gets a value indicating whether a stream supports seeking.
'
' @return <b>False</b> if the stream is closed; otherwise, <b>True</b>.
Public Property Get CanSeek() As Boolean
    CanSeek = mViewStream.CanSeek
End Property

''
' Gets a value that determines whether the current stream can time out.
'
' @return A value that determines whether the current stream can time out.
' @remarks The <b>CanTimeout</b> property always returns <b>False</b>.
Public Property Get CanTimeout() As Boolean
    CanTimeout = False
End Property

''
' Gets the stream length (size) or the total amount of memory assigned to a stream (capacity).
'
' @return The size or capacity of the stream.
' @remarks This property returns the <i>Capacity</i> value provided to the constructor. If no <i>Capacity</i> value was specified when the stream was initialized, this property returns the stream length.
Public Property Get Capacity() As Long
    Capacity = mViewStream.Capacity
End Property

''
' Gets the length of the data in a stream.
'
' @return If unchanged since the stream was initialized, this property returns the length value provided to the constructor. If changes to the stream have occurred, this property reflects any changes to the length of the data.
Public Property Get Length() As Long
    Length = mViewStream.Length
End Property

''
' Gets the current position in a stream.
'
' @return The current position in the stream.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'   </tr>
' </table>
Public Property Get Position() As Currency
    Position = mViewStream.Position
End Property

''
' Sets the current position in a stream.
'
' @param Value The current position in the stream.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td>The position is set to a negative value.</td>
'    </tr>
' </table>
Public Property Let Position(ByVal Value As Currency)
    mViewStream.Position = Value
End Property

''
' Closes the current stream and releases any resources (such as file handles) associated with the current stream.
Public Sub CloseStream()
    If Not mViewStream Is Nothing Then
        mViewStream.CloseStream
    End If
    
    ReleaseViewData
End Sub

''
' Gets a value, in miliseconds, that determines how long the stream will attempt to read before timing out.
'
' @return A value, in miliseconds, that determines how long the stream will attempt to read before timing out.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The <b>ReadTimeout</b> method always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.</td>
'   </tr>
' </table>
Public Property Get ReadTimeout() As Long
    ReadTimeout = mViewStream.ReadTimeout
End Property

''
' Sets a value, in miliseconds, that determines how long the stream will attempt to read before timing out.
'
' @param Value A value, in miliseconds, that determines how long the stream will attempt to read before timing out.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The <b>ReadTimeout</b> method always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.</td>
'   </tr>
' </table>
Public Property Let ReadTimeout(ByVal Value As Long)
    mViewStream.ReadTimeout = Value
End Property

''
' Gets a value, in miliseconds, that determines how long the stream will attempt to write before timing out.
'
' @return A value, in miliseconds, that determines how long the stream will attempt to write before timing out.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The <b>WriteTimeout</b> method always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.</td>
'   </tr>
' </table>
Public Property Get WriteTimeout() As Long
    WriteTimeout = mViewStream.WriteTimeout
End Property

''
' Sets a value, in miliseconds, that determines how long the stream will attempt to write before timing out.
'
' @param Value A value, in miliseconds, that determines how long the stream will attempt to write before timing out.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td>The <b>WriteTimeout</b> method always throws an <a href="InvalidOperationException.html">InvalidOperationException</a>.</td>
'   </tr>
' </table>
Public Property Let WriteTimeout(ByVal Value As Long)
    mViewStream.WriteTimeout = Value
End Property

''
' Begins an asynchronous read operation.
'
' @param Buffer The buffer to read the data into.
' @param Offset The byte offset in <i>buffer</i> at which to begin writing data read from the stream.
' @param Count The maximum number of bytes to read.
' @param Callback An optional asynchronous callback, to be called when the read is complete.
' @param State A user-provided object that distinguishes this particular asynchronous read request from other requests.
' @return An <a href="IAsyncResult.html">IAsyncResult</a> that represents the asynchronous read, which could still be pending.
' @remarks <p>The default implementation of <b>BeginRead</b> on a stream calls the <a href="Stream.ReadBlock.html">ReadBlock</a> method synchronously.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>One or more of the arguments is invalid.</td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>Methods were called after the stream was closed.</td>
'    </tr>
'   <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The current <b>Stream</b> implementation does not support the read operation.</td>
'   </tr>
' </table>
Public Function BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional State As Variant) As IAsyncResult
    If Not CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
        
    Set BeginRead = mViewStream.BeginRead(Buffer, Offset, Count, Callback, State)
End Function

''
' Begins an asynchronous write operation. (Consider using WriteAsync instead; see the Remarks section.)
'
' @param Buffer The buffer to write data from.
' @param Offset The byte offset in <i>Buffer</i> from which to begin writing.
' @param Count The maximum number of bytes to write.
' @param Callback An optional asynchronous callback, to be called when the write is complete.
' @param State A user-provided object that distinguishes this particular asynchronous write request from other requests.
' @return An <a href="IAsyncResult.html">IAsyncResult</a> that represents the asynchronous write, which could still be pending.
' @remarks The default implementation of <b>BeginWrite</b> on a stream calls the Write method synchronously.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>Attempted an asynchronous write past the end of the stream, or a disk error occurs.</td>
'   </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>Methods were called after the stream was closed.</td>
'    </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The current <b>Stream</b> implementation does not support the write operation.</td>
'    </tr>
' </table>
Public Function BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional State As Variant) As IAsyncResult
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
        
    Set BeginWrite = mViewStream.BeginWrite(Buffer, Offset, Count, Callback, State)
End Function

''
' Waits for the pending asynchronous read to complete.
'
' @param AsyncResult The reference to the pending asynchronous request to finish.
' @return The number of bytes read from the stream, between zero (0) and the number of bytes you requested. Streams return zero (0) only at the end of the stream, otherwise, they should block until at least one byte is available.
' @remarks <p>Call <b>EndRead</b> to determine how many bytes were read from the stream.</p>
' <p><b>EndRead</b> can be called once on every <a href="IAsyncResult.html">IAsyncResult</a> from <a href="MemoryMappedViewStream.BeginRead.html">BeginRead</a>.</p>
' <p>This method blocks until the I/O operation has completed.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>AsyncResult</i> is <b>Nothing</b>.</td>
'   </tr>
'    <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td>
'       <td>
'           A handle to the pending read operation is not available.<br/>
'           - or -<br/>
'           The pending operation does not support reading.
'       </td>
'    </tr>
'    <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td><i>AsyncResult</b> did not originate from a <a href="MemoryMappedViewStream.BeginRead.html">BeginRead</a> method on the current stream.</td>
'    </tr>
'    <tr>
'        <td><a href="IOException.html">IOException</a></td><td>The stream is closed or an internal error has occurred.</td>
'    </tr>
' </table>
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long
    EndRead = mViewStream.EndRead(AsyncResult)
End Function

''
' Ends an asynchronous write operation.
'
' @param AsyncResult A reference to the outstanding asynchronous I/O request.
' @remarks <b>EndWrite</b> must be called exactly once on every <a href="IAsyncResult.html">IAsyncResult</a> from <a href="MemoryMappedViewStream.BeginWrite.html">BeginWrite</a>.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>AsyncResult</i> is <b>Nothing</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="InvalidOperationException.html">InvalidOperationException</a></td><td><i>AsyncResult</i> did not originate from a <a href="MemoryMappedViewStream.BeginWrite.html">BeginWrite</a> method on the current stream.</td>
'    </tr>
'    <tr>
'        <td><a href="IOException.html">IOException</a></td><td>The stream is closed or an internal error has occurred./td>
'    </tr>
' </table>
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult)
    mViewStream.EndWrite AsyncResult
End Sub

''
' Clears all buffers for this stream and causes any buffered data to be written to the underlying file.
Public Sub Flush()
    FlushViewOfFile mViewAddress, mCapacity
End Sub

''
' Reads the specified number of bytes into the specified array.
'
' @param Buffer The array to read the bytes into.
' @param Offset The starting index in <i>Buffer</i> to begin reading to.
' @param Count The number of bytes to read from the stream.
' @return The number of bytes actually read from the stream.
' @return The total number of bytes read into the buffer. This can be less than the number of bytes requested if that many bytes are not currently available, or zero (0) if the end of the stream has been reached.
' @remarks The offset parameter gives the offset of the byte in the array parameter (the buffer index) at which to begin reading, and the count parameter gives the maximum number of bytes to be read from this stream. The returned value is the actual number of bytes read, or zero if the end of the stream is reached.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The <a href="MemoryMappedViewStream.get.CanRead.html">CanRead</a> property is set to <b>False</b>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>The <i>Buffer</i> parameter is set to <b>null</b>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           The <i>Offset</i> parameter is less than lower bound of <i>Buffer</i>.<br/>
'           - or -<br/>
'           The <i>Count</i> parameter is less than zero.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The <i>Offset</i> parameter plus the <i>Count</i> parameter is greater than the upper bound of the buffer array.</td>
'    </tr>
' </table>
Public Function ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    If Not CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
    
    ReadBlock = mViewStream.ReadBlock(Buffer, Offset, Count)
End Function

''
' Reads a byte from a stream and advances the position within the stream by one byte, or returns -1 if at the end of the stream.
'
' @return The unsigned byte cast to an Int32 object, or -1 if at the end of the stream.
' @remarks Use this method to return integer values from the stream.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The underlying memory does not support reading.</td>
'    </tr>
' </table>
Public Function ReadByte() As Long
    If Not CanRead Then _
        Error.NotSupported NotSupported_UnreadableStream
    
    ReadByte = mViewStream.ReadByte
End Function

''
' Sets the current position of the current stream to the given value.
'
' @param Offset The point relative to <i>Origin</i> to begin seeking from.
' @param Origin Specifies the beginning, the end, or the current position as a reference point for origin, using a value of type <a href="Stream.SeekOrigin.html">SeekOrigin</a>.
' @return The new position in the stream.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="IOException.html">IOException</a></td><td>An attempt was made to seek before the beginning of the stream.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td>The <i>Offset</i> value is larger than the maximum size of the stream.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Loc</i> is invalid.</td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'    </tr>
' </table>
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    SeekPosition = mViewStream.SeekPosition(Offset, Origin)
End Function

''
' Sets the length of the current stream.
'
' @param Value The desired length of the current stream in bytes.
' @remarks This method is internal only, because a stream that represents a view is not considered to have a length.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The method is not supported.</td>
'   </tr>
' </table>
Public Sub SetLength(ByVal Value As Currency)
    Error.NotSupported NotSupported_MMViewStreamsFixedLength
End Sub

''
' Writes a block of bytes to the current stream using data from a buffer.
'
' @param Buffer The byte array from which to copy bytes to the current stream.
' @param Offset The offset in the buffer at which to begin copying bytes to the current stream.
' @param Count The number of bytes to write to the current stream.
' @remarks Writing occurs at the current position in the stream.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The underlying memory does not support writing.<br/>
'           - or -<br/>
'           An attempt is made to write to the stream and the <a href="MemoryMappedViewStream.get.CanWrite.html">CanWrite</a> property is <b>False</b>.<br/>
'           - or -<br/>
'           The <i>Count</i> value is greater than the capacity of the stream.<br/>
'           - or -<br/>
'           The position is at the end of the stream capacity.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="IOException.html">IOException</a></td><td>An I/O error occurs.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           The <i>Offset</i> parameter is less than the lower bound of the buffer array.<br/>
'           - or -<br/>
'           The <i>Count</i> parameter is less than zero.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The <i>Offset</i> parameter plus the <i>Count</i> parameter is greater than the upper bound of the buffer array.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>The <i>Buffer</i> parameter is <b>null</b>.</td>
'    </tr>
' </table>
Public Sub WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    
    mViewStream.WriteBlock Buffer, Offset, Count
End Sub

''
' Writes a byte to the current position in the file stream.
'
' @param Value A byte value written to the stream.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The stream is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td>
'        <td>
'           The underlying memory does not support writing.<br/>
'           - or -<br/>
'           An attempt is made to write to the stream and the <a href="MemoryMappedViewStream.get.CanWrite.html">CanWrite</a> property is <b>False</b>.<br/>
'           - or -<br/>
'           The current position is at the end of the capacity of the stream.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="IOException.html">IOException</a></td><td>The supplied <i>Value</i> causes the stream exceed its maximum capacity.</td>
'    </tr>
' </table>
Public Sub WriteByte(ByVal Value As Byte)
    If Not CanWrite Then _
        Error.NotSupported NotSupported_UnwritableStream
    
    mViewStream.WriteByte Value
End Sub

''
' Reads the bytes from the current stream and writes them to another stream.
'
' @param Destination The stream to which the contents of the current stream will be copied.
' @param BufferSize The size of the buffer. This value must be greater than zero. The default size is 81920.
' @remarks Copying begins at the current position in the current stream, and does not reset the position of the destination stream after the copy operation is complete.
' <h4>Exceptions</h4>
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
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MappingHandle As SafeMemoryMappedFileHandle, ByVal Offset As Currency, ByVal Size As Long, ByVal Access As MemoryMappedFileAccess)
    Dim PageAlignedOffset As Currency
    Dim Capacity As Long
    Dim StartIndex As Long
    Dim Info As SYSTEM_INFO
    
    Set mMappingHandle = MappingHandle
    Offset = Int(Offset)
    GetSystemInfo Info
    
    ' MapViewOfFile requires the offset into a file to be virtual page aligned, so we will
    ' adjust for it and simply hide the begining of the allocated array that was not intended
    ' to be part of the view by using a starting index into the view data.
    StartIndex = CLng(Int(Offset Mod CCur(Info.dwAllocationGranularity)))
    
    If StartIndex = 0 Then
        ' the offset is already page aligned so no adjustment needed.
        PageAlignedOffset = Offset
    Else
        ' calculate the number of whole pages into the file to be used
        ' as the new offset for the view. The mStartIndex will be used to skip
        ' any portion that is in the view data that should be hidden.
        PageAlignedOffset = Int(Offset / Info.dwAllocationGranularity)
    End If
    
    With CInt64(PageAlignedOffset)
        mViewAddress = MapViewOfFile(mMappingHandle.DangerousGetHandle, GetDesiredAccess(Access), .HighPart, .LowPart, Size)
    End With
    
    If mViewAddress = vbNullPtr Then _
        Error.Win32Error Err.LastDllError
        
    SetCapacity Size
    InitViewData StartIndex
End Sub

Private Sub Class_Terminate()
    CloseStream
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCapacity(ByVal Size As Long)
    Dim Info As MEMORY_BASIC_INFORMATION
    
    If Size = 0 Then
        ' if no size is set then we use the size of the mapped region
        ' rounded up to the next virtual page boundary. For instance
        ' if the system virtual page size is 4096 bytes, then the region size
        ' will be the next even 4096 byte page boundary after the size of
        ' mapped file, so a file that is 6000 bytes will create a region
        ' size of 8192 bytes.
        VirtualQuery ByVal mViewAddress, Info, LenB(Info)
        Size = Info.RegionSize
    End If

    mCapacity = Size
End Sub

Private Sub InitViewData(ByVal StartIndex As Long)
    With mViewDataSA
        .cbElements = 1
        .cDims = 1
        .cElements = mCapacity
        .pvData = mViewAddress
    End With
    
    SAPtr(mViewData) = VarPtr(mViewDataSA)
    Set mViewStream = Cor.NewMemoryStream(mViewData, StartIndex, mCapacity - StartIndex)
End Sub

Private Sub ReleaseViewData()
    SAPtr(mViewData) = vbNullPtr
    UnmapViewOfFile mViewAddress
    mViewAddress = vbNullPtr
End Sub

Private Function GetDesiredAccess(ByVal Access As MemoryMappedFileAccess) As Long
    Select Case Access
        Case MemoryMappedFileAccess.ReadWrite
            GetDesiredAccess = FILE_MAP_ALL_ACCESS
            mCanRead = True
            mCanWrite = True
        Case MemoryMappedFileAccess.ReadOnly
            GetDesiredAccess = FILE_MAP_READ
            mCanRead = True
        Case MemoryMappedFileAccess.CopyOnWrite
            GetDesiredAccess = FILE_MAP_COPY
            mCanRead = True
            mCanWrite = True
        Case MemoryMappedFileAccess.ReadExecute
            GetDesiredAccess = FILE_MAP_READ Or FILE_MAP_EXECUTE
            mCanRead = True
        Case MemoryMappedFileAccess.ReadWriteExecute
            GetDesiredAccess = FILE_MAP_ALL_ACCESS Or FILE_MAP_EXECUTE
            mCanRead = True
            mCanWrite = True
        Case MemoryMappedFileAccess.WriteOnly
            GetDesiredAccess = FILE_MAP_WRITE
            mCanWrite = True
        Case Else
            Error.ArgumentOutOfRange "Access"
    End Select
End Function


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
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional State As Variant) As IAsyncResult
    Set Stream_BeginRead = BeginRead(Buffer, Offset, Count, Callback, State)
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional State As Variant) As IAsyncResult
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

Private Property Let Stream_Position(ByVal RHS As Currency)
    Position = RHS
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

Private Property Let Stream_ReadTimeout(ByVal RHS As Long)
    ReadTimeout = RHS
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

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)
    WriteTimeout = RHS
End Property

Private Property Get Stream_WriteTimeout() As Long
    Stream_WriteTimeout = WriteTimeout
End Property
