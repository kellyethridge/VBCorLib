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
'    Module: CryptoStream
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
    Throw Cor.NewNotSupportedException("Seeking is not supported.")
End Property

''
' Sets the current position within the Stream.
'
' @param RHS The new position.
' @remarks Always throws a NotSupportedException exception.
'
Public Property Let Position(ByVal RHS As Currency)
    Throw Cor.NewNotSupportedException("Seeking is not supported.")
End Property

''
' Returns the length of the current Stream.
'
' @return The length of the Stream.
' @remarks Always throws a NotSupportedException exception.
'
Public Property Get Length() As Currency
    Throw Cor.NewNotSupportedException("Seeking is not supported.")
End Property

''
' Sets the length of the current Stream.
'
' @param Value The new length of the Stream.
' @remarks Always throws a NotSupportedException exception.
'
Public Sub SetLength(ByVal Value As Currency)
    Throw Cor.NewNotSupportedException("Seeking is not supported.")
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
    Throw Cor.NewNotSupportedException("Seeking is not supported.")
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
    If Me.CanWrite And (Not mFinalTransformed) Then Call FlushFinalBlock
    Call mStream.CloseStream
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
    Call mStream.Flush
End Sub

''
' Flushes the remaining data to the underlying stream.
'
' @remarks This can only be called once.
'
Public Sub FlushFinalBlock()
    If Not CanWrite Then _
        Throw Cor.NewNotSupportedException("Stream does not support writing.")
    If mFinalTransformed Then _
        Throw Cor.NewNotSupportedException("Cannot call FlushFinalBlock twice.")
    
    mOutputBuffer = mTransform.TransformFinalBlock(mInputBuffer, 0, mBytesBuffered)
    Call mStream.WriteBlock(mOutputBuffer, 0, CorArray.Length(mOutputBuffer))
    
    If TypeOf mStream Is CryptoStream Then
        Dim cs As CryptoStream
        Set cs = mStream
        Call cs.FlushFinalBlock
    End If
    
    Call Flush
    Call InitBuffers
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
        Throw Cor.NewNotSupportedException("The Stream does not support reading.")
    If CorArray.IsNull(Buffer) Then _
        Throw Cor.NewArgumentNullException("Buffer", Environment.GetResourceString(ArgumentNull_Array))
    If Offset < LBound(Buffer) Then _
        Throw Cor.NewArgumentOutOfRangeException("Offset", Message:=Environment.GetResourceString(ArgumentOutOfRange_LBound))
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
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
        Throw Cor.NewNotSupportedException("The Stream does not support reading.")
    
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
        Throw Cor.NewNotSupportedException("The Stream does not support reading.")
    
    Dim Ret As StreamAsyncResult
    Set Ret = Cor.NewStreamAsyncResult(State)
    
    On Error GoTo errTrap
    With Ret
        .BytesRead = ReadBlock(Buffer, Offset, Count)
        .IsCompleted = True
        .IsReadType = True
    End With

    If Not Callback Is Nothing Then Call Callback.Execute(Ret)

errTrap:
    Dim Ex As Exception
    If Catch(Ex, Err) Then Set Ret.Exception = Ex
    Set BeginRead = Ret
End Function

''
' Signifies the end of an asynchronous read from the stream.
'
' @param AsyncResult The IAsyncResult object returned from the BeginRead function.
' @return The number of bytes read from the stream in to <i>Buffer</i> in the BeginRead method.
'
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    
    Dim Result As StreamAsyncResult
    Set Result = AsyncResult
    If Not Result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If Result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")
    
    If Not Result.Exception Is Nothing Then Throw Result.Exception
    
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
        Throw Cor.NewNotSupportedException("Stream does not support writing.")
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Buffer), Offset, Count)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Buffer", Offset, "Offset", Count, "Count", False)
    
    If (Count > mInputBlockSize) And mTransform.CanTransformMultipleBlocks Then
        Call WriteMultiBlock(Buffer, Offset, Count)
    End If
    
    Do While Count > 0
        Call WriteToBuffer(Buffer, Offset, Count)
    Loop
End Sub

''
' Writes a single byte to the Stream.
'
' @param Value The byte to be written to the stream.
'
Public Sub WriteByte(ByVal Value As Byte)
    If Not CanWrite Then _
        Throw Cor.NewNotSupportedException("Stream does not support writing.")
    
    mInputBuffer(mInputIndex) = Value
    mInputIndex = mInputIndex + 1
    mBytesBuffered = mBytesBuffered + 1
    
    If mBytesBuffered = mInputBlockSize Then Call FlushBuffer
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
    Dim Ret As StreamAsyncResult
    
    Set Ret = Cor.NewStreamAsyncResult(State)
    On Error GoTo errTrap:
    With Ret
        .CompletedSynchronously = True
        .IsCompleted = True
    End With
    Call WriteBlock(Buffer, Offset, Count)
    
    If Not Callback Is Nothing Then Call Callback.Execute(Ret)
    
errTrap:
    Dim Ex As Exception
    If Catch(Ex, Err) Then Set Ret.Exception = Ex
    Set BeginWrite = Ret
End Function

''
' Signifies the end of an asynchronous write to the stream.
'
' @param AsyncResult The IAsyncResult returned from the BeginWrite function.
'
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult)
    If AsyncResult Is Nothing Then _
        Throw Cor.NewArgumentNullException("AsyncResult")
    If Not TypeOf AsyncResult Is StreamAsyncResult Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    
    Dim Result As StreamAsyncResult
    Set Result = AsyncResult
    If Result.IsReadType Then _
        Throw Cor.NewArgumentException("AsyncResult object is not part of this stream.", "AsyncResult")
    If Result.EndCalled Then _
        Throw Cor.NewInvalidOperationException("The EndRead has already been called.")

    If Not Result.Exception Is Nothing Then Throw Result.Exception
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Stream As Stream, ByVal Transform As ICryptoTransform, ByVal Mode As CryptoStreamMode)
    If Stream Is Nothing Then _
        Throw Cor.NewArgumentNullException("Stream", Environment.GetResourceString(ArgumentNull_Stream))
    If Transform Is Nothing Then _
        Throw Cor.NewArgumentNullException("Transform")
    
    Select Case Mode
        Case CryptoStreamMode.ReadMode
            If Not Stream.CanRead Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_StreamNotReadable), "Stream")
                
        Case CryptoStreamMode.WriteMode
            If Not Stream.CanWrite Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_StreamNotWritable), "Stream")
                
        Case Else
            Throw Cor.NewArgumentException("Invalid mode.", "mode")
    End Select
    
    Set mStream = Stream
    Set mTransform = Transform
    mMode = Mode
    
    mOutputBlockSize = mTransform.OutputBlockSize
    mInputBlockSize = mTransform.InputBlockSize
    
    Call InitBuffers
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
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
    Call mStream.WriteBlock(mOutputBuffer, 0, TotalBytesTransformed)
    
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
    If AmountToCopy = 0 Then Exit Sub
    
    ' Fill our buffer with as much as possible. We still may not have filled it.
    Call CopyMemory(mInputBuffer(mInputIndex), Buffer(Offset), AmountToCopy)
    
    ' And we update everyone according to how much we moved.
    mBytesBuffered = mBytesBuffered + AmountToCopy
    mInputIndex = mInputIndex + AmountToCopy
    Offset = Offset + AmountToCopy
    Count = Count - AmountToCopy
    
    ' And if we did fill the buffer, get it transformed and written to the Stream.
    If mBytesBuffered = mInputBlockSize Then Call FlushBuffer
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
    If mBytesBuffered > 0 Then Call WriteToBuffer(InputBuffer, InputOffset, InputCount)
    
    ' If there isn't even enough bytes to fill a block, then how can we do multiblock?
    ' InputCount can be modified in the WriteToBuffer call.
    If InputCount <= mInputBlockSize Then Exit Sub
    
    ' How many blocks will we be doing at once?
    Dim TotalBytes As Long
    TotalBytes = (InputCount \ mOutputBlockSize) * mInputBlockSize
        
    Dim OutputBuffer() As Byte
    ReDim OutputBuffer(0 To TotalBytes - 1)
    
    ' apply the cipher to the big array of bytes.
    Dim TotalBytesTransformed As Long
    TotalBytesTransformed = mTransform.TransformBlock(InputBuffer, InputOffset, TotalBytes, OutputBuffer, 0)
    
    ' If anything was transformed, then write it out to the Stream.
    If TotalBytesTransformed > 0 Then
        Call mStream.WriteBlock(OutputBuffer, 0, TotalBytesTransformed)
        InputOffset = InputOffset + TotalBytes
        InputCount = InputCount - TotalBytes
    End If
End Sub

Private Function ReadBufferedBytes(ByRef Buffer() As Byte, ByRef Offset As Long, ByRef Count As Long) As Long
    ' If nothing is buffered, we can't read from it
    ' and return what it contains.
    If mBytesBuffered = 0 Then Exit Function
    
    Dim AmountToCopy As Long
    
    ' Only copy what we want or what we already have, which
    ' ever is the smaller value.
    AmountToCopy = CorMath.Min(mBytesBuffered, Count)
    
    ' Make sure there is room in the return buffer.
    If (Offset + AmountToCopy - 1) > UBound(Buffer) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Buffer")
    
    ' Copy the data to the return buffer and update all the counters.
    Call CopyMemory(Buffer(Offset), mOutputBuffer(mOutputIndex), AmountToCopy)
    Offset = Offset + AmountToCopy
    Count = Count - AmountToCopy
    mBytesBuffered = mBytesBuffered - AmountToCopy
    mOutputIndex = mOutputIndex + AmountToCopy
    ReadBufferedBytes = AmountToCopy
End Function

Private Function ReadMultiBlock(ByRef Buffer() As Byte, ByRef Offset As Long, ByRef Count As Long) As Long
    ' If less than atleast 1 full block is requested, then
    ' we can't perform a mutli-block read, so exit.
    If Count < mOutputBlockSize Then Exit Function
    
    Dim Result As Long
    
    ' Calculate how many full output blocks we want, then determine
    ' how many input blocks in bytes it will take to get it.
    Dim TotalBytes As Long
    TotalBytes = (Count \ mOutputBlockSize) * mInputBlockSize
    
    Dim InputBuffer() As Byte
    ReDim InputBuffer(0 To TotalBytes - 1)
    
    Dim BytesRead As Long
    BytesRead = mStream.ReadBlock(InputBuffer, 0, TotalBytes)
    
    If BytesRead = 0 Then Exit Function
    
    Dim TotalBytesTransformed As Long
    TotalBytesTransformed = mTransform.TransformBlock(InputBuffer, 0, BytesRead, Buffer, Offset)
    
    Offset = Offset + TotalBytesTransformed
    Count = Count - TotalBytesTransformed
    Result = Result + TotalBytesTransformed
    
    ReadMultiBlock = Result
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call CloseStream
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
'   Stream Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Stream_BeginRead = BeginRead(Buffer, Offset, Count, Callback, State)
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
    Call BeginWrite(Buffer, Offset, Count, Callback, State)
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
    Call CloseStream
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    Stream_EndRead = EndRead(AsyncResult)
End Function

Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    Call EndWrite(AsyncResult)
End Sub

Private Function Stream_Equals(Value As Variant) As Boolean
    Stream_Equals = Equals(Value)
End Function

Private Sub Stream_Flush()
    Call Flush
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
    Call SetLength(Value)
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = ToString
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    Call WriteBlock(Buffer, Offset, Count)
End Sub

Private Sub Stream_WriteByte(ByVal Value As Byte)
    Call WriteByte(Value)
End Sub

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)
    WriteTimeout = RHS
End Property

Private Property Get Stream_WriteTimeout() As Long
    Stream_WriteTimeout = WriteTimeout
End Property
