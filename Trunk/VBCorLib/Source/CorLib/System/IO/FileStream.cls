VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: FileStream
'

''
' Represents a file as a Stream.
'
' @remarks A <b>FileStream</b> object cannot be created directly. In order to create a new
' <b>FileStream</b> object, use the <b>Cor.NewFileStream</b> method.
' <p>The FileStream can be used to create and modify files saved to disk. Through the stream
' the file can be navigated to specified byte positions within the file and the data can be read
' or written. Writing data will overwrite the data currently in the file at that location. As
' data is read or written, the file pointer is advanced the number of bytes read or written.</p>
' <p>Additionally, and existing handle to a file, pipe, or mailslot can have a stream created
' around it. These streams are not seekable, meaning the position within the stream cannot be
' set manually. Reading and writing to these streams still cause the pointer to advance the
' appropriate number of bytes.</p>
' <p>The FileStream object buffers reads and writes to the underlying stream for improved
' performance. It is assumed that several reads or several writes will occur repeatedly.
' This allows a single buffer to be used between reading and writing. If reading begins and
' there is data written, the data is flushed out to the stream to ensure the stream remains
' in sync. The buffer is then filled with <i>BufferSize</i> bytes for a quicker read.</p>
'
' @see Constructors
' @see Stream
' @include "..\..\Includes\FileStream.txt"
Option Explicit
Implements IObject
Implements Stream

Private Const MinBufferSize            As Long = 8

'Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
'Private Const FILE_TYPE_DISK            As Long = 1
'Private Const INVALID_SET_FILE_POINTER  As Long = -1
'Private Const INVALID_FILE_SIZE         As Long = -1
'Private Const ERROR_BROKEN_PIPE         As Long = 109

'Public Type NativeOverlapped
'    InternalLow     As Long
'    InternalHigh    As Long
'    OffsetLow       As Long
'    OffsetHigh      As Long
'    EventHandle     As Long
'    CallbackHandle  As Long
'End Type



''
' The modes in which to open a FileStream
'
' @param Append Opens an existing or creates a new file, then moves to the end
' for writing to begin at. Reading of the stream is not supported.
' @param Create Creates a new file. If the file already exists, it is overwritten.
' @param CreateNew Creates a new file. If the file already exists, then an IOException is thrown.
' @param OpenExisting Opens an existing file. If the file does not exist, then a FileNotFoundException is thrown.
' @param OpenOrCreate Opens an existing file. If the file does not exist, it will be created.
' @param Truncate Opens an existing file and sets the file length to 0. If an attempt to
' read is made, then an exception will be thrown.
'
Public Enum FileMode
    Append = &H80
    Create = 2
    CreateNew = 1
    OpenExisting = 3
    OpenOrCreate = 4
    Truncate = 5
End Enum

''
' The accesses allowed to the FileStream.
'
' @param ReadAccess Grants read permission to the FileStream.
' @param WriteAccess Grants write permission to the FileStream.
' @param ReadWriteAccess Grants both read and write permissions to the FileStream.
'
Public Enum FileAccess
    DefaultAccess = 0
    ReadAccess = &H80000000
    WriteAccess = &H40000000
    ReadWriteAccess = &HC0000000
End Enum

''
' How the underlying stream can be shared with other processes.
'
' @param None The stream is not shared with other processes.
' @param ReadShare Other processes can access the file with read access only.
' @param WriteShare Other processes can access the file with write access only.
' @param ReadWriteShare Other processes can access the file with read and write access.
'
Public Enum FileShare
    None = 0
    ReadShare = 1
    WriteShare = 2
    ReadWriteShare = 3
End Enum

Private mName           As String
Private mHandle         As SafeFileHandle
Private mCanRead        As Boolean
Private mCanWrite       As Boolean
Private mCanSeek        As Boolean
Private mMinPosition    As Currency
Private mFilePosition   As Currency
Private mWritePosition  As Long
Private mReadPosition   As Long
Private mBuffer()       As Byte
Private mBufferSize     As Long
Private mBufferOffset   As Currency
Private mBytesBuffered  As Long
Private mIsAsync        As Boolean
Private mActiveAsyncResult As IAsyncResult


''
' Returns if this object can timeout.
'
' @return Always returns False for a FileStream.
'
Public Property Get CanTimeout() As Boolean
    CanTimeout = False
End Property

''
' Returns the read timeout duration.
'
' @return The timeout duration.
' @remarks This is not supported in FileStream. An exception will be thrown.
'
Public Property Get ReadTimeout() As Long
    Throw Error.InvalidOperation(InvalidOperation_Timeouts)
End Property

''
' Sets the read timeout duration.
'
' @param RHS The timeout duration in milliseconds.
' @remarks This is not supported in FileStream. An exception will be thrown.
'
Public Property Let ReadTimeout(ByVal Value As Long)
    Throw Error.InvalidOperation(InvalidOperation_Timeouts)
End Property

''
' Returns the write timeout duration.
'
' @return The write timeout duration.
' @remarks This is not supported in FileStream. An exception will be thrown.
'
Public Property Get WriteTimeout() As Long
    Throw Error.InvalidOperation(InvalidOperation_Timeouts)
End Property

''
' Sets the write timeout duration.
'
' @param RHS The timeout duration in milliseconds.
' @remarks This is not supported in FileStream. An exception will be thrown.
'
Public Property Let WriteTimeout(ByVal Value As Long)
    Throw Error.InvalidOperation(InvalidOperation_Timeouts)
End Property

''
' Returns whether or not the FileStream was opened up in asynchronous mode.
'
' @return Boolean indicating if the FileStream is in asynchronous mode or not.
'
Public Property Get IsAsync() As Boolean
    IsAsync = mIsAsync
End Property

''
' Returns the name of the underlying stream.
'
' @return If the stream is a file, then the file name is returned, otherwise the name is unknown.
'
Public Property Get Name() As String
    If Len(mName) = 0 Then
        Name = "Unknown Stream Name"
    Else
        Name = mName
    End If
End Property

''
' Returns an indication for the ability to read from the file stream.
'
' @return An indication for being able to read from the file stream.
'
Public Property Get CanRead() As Boolean
    CanRead = (mCanRead And Not mHandle.IsClosed)
End Property

''
' Returns an indication for the ability to write to the file stream.
'
' @return An indication for being able to write to the file stream.

Public Property Get CanWrite() As Boolean
    CanWrite = (mCanWrite And Not mHandle.IsClosed)
End Property

''
' Returns an indication for the ability to seek within the file stream.
'
' @return An indication for being able to seek within the file stream.
'
Public Property Get CanSeek() As Boolean
    CanSeek = (mCanSeek And Not mHandle.IsClosed)
End Property

''
' Returns the handle to the underlying stream.
'
' @return The handle for the underlying stream.
' @remarks If there is any information in the buffers, they are flushed
' back to the underlying stream before the handle is returned.
'
Public Property Get SafeFileHandle() As SafeFileHandle
    If mCanWrite Then
        Flush
    End If
    mReadPosition = 0
    mBytesBuffered = 0
    Set SafeFileHandle = mHandle
End Property

''
' Locks a portion of a file to prevent write access for other processes.
'
' @param Position The starting position in the file to begin locking.
' @param Length The number of bytes to be locked.
' @remarks Onces a section of the file is locked, read access is still allowed to other processes.
'
Public Sub LockStream(ByVal Position As Currency, ByVal Length As Currency)
    VerifyIsOpen
    
    Position = Int(Position)
    If Position < 0@ Then _
        Throw Error.ArgumentOutOfRange("Position", ArgumentOutOfRange_NeedNonNegNum)
    
    Length = Int(Length)
    If Length < 0@ Then _
        Throw Error.ArgumentOutOfRange("Length", ArgumentOutOfRange_NeedNonNegNum)
    
    Dim pos As DLong
    AsCurr(pos) = CCur(0.0001@ * Position)
    With AsDLong(CCur(0.0001@ * Length))
        If LockFile(mHandle, pos.LoDWord, pos.HiDWord, .LoDWord, .HiDWord) = BOOL_FALSE Then
            IOError.WinIOError Err.LastDllError
        End If
    End With
End Sub

''
' Unlocks a portion of the file to allow write access to the stream for other processes.
'
' @param Position The starting position in the file to begin unlocking.
' @param Length The number of bytes to be unlocked.
'
Public Sub UnlockStream(ByVal Position As Currency, ByVal Length As Currency)
    VerifyIsOpen
    
    Position = Int(Position)
    If Position < 0@ Then _
        Throw Error.ArgumentOutOfRange("Position", ArgumentOutOfRange_NeedNonNegNum)
    
    Length = Int(Length)
    If Length < 0@ Then _
        Throw Error.ArgumentOutOfRange("Length", ArgumentOutOfRange_NeedNonNegNum)
    
    Dim pos As DLong
    AsCurr(pos) = CCur(0.0001@ * Position)
    With AsDLong(CCur(0.0001@ * Length))
        If UnlockFile(mHandle, pos.LoDWord, pos.HiDWord, .LoDWord, .HiDWord) = BOOL_FALSE Then
            IOError.WinIOError Err.LastDllError
        End If
    End With
End Sub

''
' The length of the current stream in bytes.
'
' @return The number of bytes the current file is.
' @remarks If there is write data in the buffer, it will be
' flushed out to the file to allow for an accurate reading.
'
Public Property Get Length() As Currency
    VerifyCanSeek
    VerifyIsOpen
    
    If mCanWrite Then
        Flush
    End If
    
    Dim Size As DLong
    With Size
        .LoDWord = GetFileSize(mHandle, .HiDWord)
        ' in order to catch an error we check the return value against
        ' INVALID_FILE_SIZE (&HFFFFFFFF). Even though it is a value
        ' of -1, the value must be treated as unsigned, so its negative
        ' values are = 2gig to 4gig positive values, which works with the
        ' 64bit file lengths. So, if we get the -1, we need to check for
        ' an actual error using Err.LastDllError.
        If .LoDWord = INVALID_FILE_SIZE Then
            Dim DllError As Long
            DllError = Err.LastDllError
            If DllError <> NO_ERROR Then
                IOError.WinIOError DllError
            End If
        End If
    End With
    Length = 10000@ * AsCurr(Size)
End Property

''
' Returns the current position within the stream.
'
' @return The current position within the stream.
' @remarks The current position is where the next byte is read from
' or written to. Once a byte has been read or written, the position
' is advanced byte 1 byte.
'
Public Property Get Position() As Currency
    Position = mFilePosition
End Property

''
' Sets the current position within the stream to be read from or written to.
'
' @param Value The new position in the stream.
' @remarks The position is an absolute byte position from the beginning of
' the file. If the new position is negative an exception is thrown. If the
' FileStream is opened in Append mode, then the position cannot be set to
' data within the original file.
'
Public Property Let Position(ByVal Value As Currency)
    If Value < 0@ Then _
        Throw Error.ArgumentOutOfRange("Value", ArgumentOutOfRange_NeedNonNegNum)
        
    SeekPosition Value, SeekOrigin.FromBeginning
End Property

''
' Moves the file pointer to a new position relative to a specified reference.
'
' @param Offset The number of bytes to move from the reference point. The value can be negative to move backwards.
' @param Origin The reference point to begin moving the file pointer from.
' @return The new position in the file stream.
'
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    VerifyCanSeek
    VerifyIsOpen
    
    Offset = Int(Offset)
    
    Dim NewPosition As Currency
    Select Case Origin
        Case FromBeginning:     NewPosition = Offset
        Case FromCurrent:       NewPosition = mFilePosition + Offset
        Case FromEnd:           NewPosition = Length + Offset
        Case Else
            Throw Error.Argument(Argument_InvalidSeekOrigin)
    End Select
    
    If NewPosition < mMinPosition Then
        If mMinPosition = 0@ Then
            Throw Error.Argument(Argument_InvalidSeekOffset)
        Else
            Throw Cor.NewIOException("Seeking to a position prior to append-start position may cause data to be overwritten.")
        End If
    End If
    
    If mCanWrite Then
        Flush
    End If
    
    Select Case NewPosition
        Case mBufferOffset To mBufferOffset + mBytesBuffered - 1
            ' in same buffer works only for reading
            mReadPosition = NewPosition - mBufferOffset
            mFilePosition = NewPosition
        Case Else
            mFilePosition = InternalSeek(Offset, Origin)
            mBufferOffset = mFilePosition
            mBytesBuffered = 0
    End Select
    
    SeekPosition = mFilePosition
End Function

''
' Sets the length of the stream to the specified length.
'
' @param Value The length of the new file in bytes.
' @remarks <p>If the new value is less than the origin file, then the file will
' be truncated. If the new value is larger than the original, then the newly
' exposed bytes are of undefined value.</p>
' <p>The stream must have Seek and Write support</p>
'
Public Sub SetLength(ByVal Value As Currency)
    VerifyCanSeek
    VerifyCanWrite
    VerifyIsOpen
    
    Value = Int(Value)
    If Value < mMinPosition Then
        If mMinPosition = 0@ Then
            Throw Error.ArgumentOutOfRange("Value", ArgumentOutOfRange_NeedNonNegNum)
        Else
            Throw Cor.NewIOException("Setting the length to a position prior to the append-start postion is not permitted.")
        End If
    End If

    Flush
    mBytesBuffered = 0
    
    InternalSeek Value, FromBeginning
    If SetEndOfFile(mHandle) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
End Sub

''
' Returns the next byte in the file stream starting at the current file position.
'
' @return The byte value at the current file position.
' @remarks If the file position is passed the end of the stream, then -1 is returned.
'
Public Function ReadByte() As Long
    VerifyCanRead
    VerifyIsOpen

    If mCanWrite Then
        Flush
    End If
    
    If Position >= Length Then
        ReadByte = -1
        Exit Function
    End If

    If (mBytesBuffered = 0) Or (mReadPosition = mBytesBuffered) Then
        FillBuffer
    End If
    
    ReadByte = mBuffer(mReadPosition)
    mReadPosition = mReadPosition + 1
    mFilePosition = mFilePosition + 1
End Function

''
' Reads a specified number of bytes into the given array.
'
' @param Bytes The array to store the bytes that are read from the stream.
' @param Offset The index in <i>Buffer</i> to begin storing bytes. Offset starts at LBound of Bytes.
' @param Count The number of bytes to be read from the stream.
' @return The number of byte actually read from the stream.
'
Public Function ReadBlock(ByRef Bytes() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    VerifyCanRead
    VerifyIsOpen
    CheckValidSingleDimArray Bytes, Parameter_Bytes
    
    If Offset < LBound(Bytes) Then _
        Throw Error.ArgumentOutOfRange("Offset", ArgumentOutOfRange_LBound)
    If Count < 0 Then _
        Throw Error.ArgumentOutOfRange("Count", ArgumentOutOfRange_NeedNonNegNum)
    If Offset + Count - 1 > UBound(Bytes) Then _
        Throw Error.Argument(Argument_InvalidOffLen)
    
    If Count = 0 Then
        Exit Function
    End If
    
    If mCanSeek Then
        If mFilePosition > Length Then
            Exit Function
        End If
    End If
    
    If (mBytesBuffered = 0) Or (mReadPosition = mBytesBuffered) Then
        If mCanWrite And mWritePosition > 0 Then
            Flush
        End If
        
        FillBuffer
    End If
    
    If mBytesBuffered = 0 Then
        Exit Function
    End If
    
    Dim Available As Long
    Available = mBytesBuffered - mReadPosition
    
    If Count > Available Then
        CopyMemory Bytes(Offset), mBuffer(mReadPosition), Available
        
        ' Read the rest of the requested amount of data from the file.
        If ReadFile(mHandle, Bytes(Offset + Available), Count - Available, Count, ByVal 0&) = BOOL_FALSE Then
            If mHandle <> ERROR_BROKEN_PIPE Then
                IOError.WinIOError Err.LastDllError
            End If
        End If
        
        Count = Count + Available
        mBytesBuffered = 0
    Else
        ' All of the data was already cached in the buffer.
        CopyMemory Bytes(Offset), mBuffer(mReadPosition), Count
    End If
    
    mReadPosition = mReadPosition + Count
    mFilePosition = mFilePosition + Count
    ReadBlock = Count
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
    Dim Ret As StreamAsyncResult
    Set Ret = Cor.NewStreamAsyncResult(State)
    
    On Error GoTo Catch
    With Ret
        .BytesRead = ReadBlock(Buffer, Offset, Count)
        .IsCompleted = True
        .IsReadType = True
    End With

    If Not Callback Is Nothing Then
        Callback.Execute Ret
    End If
    
    GoTo Finally
    
Catch:
    Dim Ex As Exception
    If Catch(Ex, Err) Then
        Set Ret.Exception = Ex
    End If
    
Finally:
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
        Throw Error.ArgumentNull("AsyncResult")
    If mActiveAsyncResult Is Nothing Then _
        Throw Error.Argument(InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple, "AsyncResult")
    If Not AsyncResult Is mActiveAsyncResult Then _
        Throw Error.InvalidOperation(InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple)
    
    Dim Result As StreamAsyncResult
    Set Result = AsyncResult
    If Not Result.IsReadType Then _
        Throw Error.Argument(InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple, "AsyncResult")
    If Result.EndCalled Then _
        Throw Error.Argument(InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple, "AsyncResult")
    If Not Result.Exception Is Nothing Then _
        Throw Result.Exception
    
    Set mActiveAsyncResult = Nothing
    
    EndRead = Result.BytesRead
End Function

''
' Writes a single byte to the stream.
'
' @param Value The value to be written to the stream.
' @remarks The value is written to the underlying buffer. Once the buffer is full, it is
' then flushed to the underlying stream and emptied to be reused.
'
Public Sub WriteByte(ByVal Value As Byte)
    VerifyCanWrite
    VerifyIsOpen
    
    If mBytesBuffered > 0 Then
        InternalSeek mFilePosition, FromBeginning
    End If
    If mWritePosition = mBufferSize Then
        Flush
    End If
    If SAPtr(mBuffer) = vbNullPtr Then
        ReDim mBuffer(0 To mBufferSize - 1)
    End If
    
    mBuffer(mWritePosition) = Value
    mWritePosition = mWritePosition + 1
    mFilePosition = mFilePosition + 1
End Sub

''
' Writes an array of bytes to the stream.
'
' @param Bytes The source array to write to the stream.
' @param Offset The index of the first byte in <i>Buffer</i> to be written to the stream.
' @param Count The number of bytes to write from the array.
' @remarks The data to be written is buffered if there is enough room, otherwise,
' any data in the buffer is written to the stream, followed by the data in <i>Buffer</i>.
'
Public Sub WriteBlock(ByRef Bytes() As Byte, ByVal Offset As Long, ByVal Count As Long)
    VerifyCanWrite
    VerifyIsOpen
    CheckValidSingleDimArray Bytes, Parameter_Bytes
    
    If Offset < LBound(Bytes) Then _
        Throw Error.ArgumentOutOfRange("Offset", ArgumentOutOfRange_LBound)
    If Count < 0 Then _
        Throw Error.ArgumentOutOfRange("Count", ArgumentOutOfRange_NeedNonNegNum)
    If Offset + Count - 1 > UBound(Bytes) Then _
        Throw Error.Argument(Argument_InvalidOffLen)
    
    If Count = 0 Then
        Exit Sub
    End If
    If SAPtr(mBuffer) = vbNullPtr Then
        ReDim mBuffer(0 To mBufferSize - 1)
    End If
    
    If mBytesBuffered > 0 Then
        InternalSeek mFilePosition, FromBeginning
    End If
    
    ' If we can't fit the data in the remaining available
    ' space in the buffer, then flush the buffer.
    If Count > mBufferSize - mWritePosition Then
        Flush
    End If
    
    ' If more bytes are to be written than can be held
    ' in the buffer, then just write it all out to the file.
    If Count > mBufferSize Then
        Dim BytesWritten As Long
        If WriteFile(mHandle, Bytes(Offset), Count, BytesWritten, ByVal 0&) = BOOL_FALSE Then
            IOError.WinIOError Err.LastDllError
        End If
        
        Count = BytesWritten
        mWritePosition = 0
    Else
        ' Cache it to the local buffer.
        CopyMemory mBuffer(mWritePosition), Bytes(Offset), Count
        mWritePosition = mWritePosition + Count
    End If
    
    mFilePosition = mFilePosition + Count
End Sub

''
' Begins an asynchronous buffer write. Currently the FileStream class does not
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
    
    On Error GoTo Catch
    With Ret
        .CompletedSynchronously = True
        .IsCompleted = True
    End With
    
    WriteBlock Buffer, Offset, Count
    
    If Not Callback Is Nothing Then
        Callback.Execute Ret
    End If
    
    GoTo Finally
    
Catch:
    Dim Ex As Exception
    If Catch(Ex, Err) Then
        Set Ret.Exception = Ex
    End If
    
Finally:
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
' Writes any data that may be in the write buffer to the underlying stream.
'
Public Sub Flush()
    VerifyCanWrite
    VerifyIsOpen
    
    Dim BytesWritten As Long
    If mWritePosition > 0 Then
        If WriteFile(mHandle, mBuffer(0), mWritePosition, BytesWritten, ByVal 0&) = BOOL_FALSE Then
            IOError.WinIOError Err.LastDllError
        End If
        
        mWritePosition = 0
        mBytesBuffered = 0
        mBufferOffset = mBufferOffset + BytesWritten
    End If
End Sub

''
' Closes the current stream, flushing any data that may need to be written to the stream.
'
' @remarks If the stream was created around a handle and the stream was given ownership
' of the handle, then the handle will also be closed with the stream. Otherwise, only
' the stream will be closed, leaving the handle in its original state.
'
Public Sub CloseStream()
    If mHandle Is Nothing Then
        Exit Sub
    End If
    
    If Not mHandle.IsClosed Then
        If mCanWrite Then
            Flush
        End If
        mHandle.CloseHandle
    End If
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Path As String, ByVal Mode As FileMode, ByVal Access As FileAccess, ByVal Share As FileShare, ByVal BufferSize As Long, ByVal UseAsync As Boolean)
    Path = CorString.Trim(Path)
    If Len(Path) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "Path")
    If BufferSize <= 0 Then _
        Throw Error.ArgumentOutOfRange("BufferSize", ArgumentOutOfRange_NeedPosNum)
    
    ValidateFileMode Mode
    ValidateFileAccess Access
    ValidateFileShare Share
    
    Dim AppendMode As Boolean
    If Mode = FileMode.Append Then
        If Access = DefaultAccess Then
            Access = WriteAccess
        ElseIf Access And ReadAccess Then
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidFileModeAndAccessCombo, GetFileModeDisplayName(Mode), GetFileAccessDisplayName(Access)))
        End If
        AppendMode = True
        Mode = FileMode.OpenOrCreate
    End If
    
    If Access = DefaultAccess Then
        Access = ReadWriteAccess
    End If
    
    Statics.Path.CheckValidPathAndFixSlashes Path
    
    Set mHandle = Win32Native.SafeCreateFile(Path, Access, Share, Mode)
    If mHandle.IsInvalid Then
        IOError.WinIOError Err.LastDllError, Path
    End If
    
    mName = Statics.Path.GetFileName(Path)
    mCanWrite = CBool(Access And FileAccess.WriteAccess)
    mCanRead = CBool(Access And FileAccess.ReadAccess) And (Mode <> Truncate)
    mCanSeek = (GetFileType(mHandle) = FILE_TYPE_DISK)
    
    If AppendMode And mCanSeek Then
        mMinPosition = InternalSeek(0@, FromEnd)
        mFilePosition = mMinPosition
        mBufferOffset = mMinPosition
    End If
    
    mBufferSize = MathExt.Max(BufferSize, MinBufferSize)
End Sub

Friend Sub InitWithHandle(ByVal Handle As SafeFileHandle, ByVal Access As FileAccess, ByVal BufferSize As Long)
    If Handle.IsInvalid Then _
        Throw Error.Argument(Argument_InvalidHandle)
    If BufferSize < 1 Then _
        Throw Error.ArgumentOutOfRange("BufferSize", ArgumentOutOfRange_NeedNonNegNum)
    
    Set mHandle = Handle
    mBufferSize = MathExt.Max(BufferSize, MinBufferSize)
    mCanSeek = (GetFileType(Handle) = FILE_TYPE_DISK)
    mCanWrite = CBool(Access And WriteAccess)
    mCanRead = CBool(Access And ReadAccess)
    
    If mCanSeek Then
        mFilePosition = InternalSeek(0@, FromCurrent)
        If mFilePosition > Length Then
            mFilePosition = InternalSeek(0@, FromEnd)
        End If
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyCanSeek()
    If Not mCanSeek Then
        Error.NotSupported NotSupported_UnseekableStream
    End If
End Sub

Private Sub VerifyIsOpen()
    If mHandle.IsClosed Then
        Throw Cor.NewObjectDisposedException("Stream", Environment.GetResourceString(ObjectDisposed_StreamClosed))
    End If
End Sub

Private Sub VerifyCanWrite()
    If Not mCanWrite Then
        Error.NotSupported NotSupported_UnwritableStream
    End If
End Sub

Private Sub VerifyCanRead()
    If Not mCanRead Then
        Error.NotSupported NotSupported_UnreadableStream
    End If
End Sub

Private Sub FillBuffer()
    mBufferOffset = mFilePosition
    mReadPosition = 0
    mWritePosition = 0
    
    If SAPtr(mBuffer) = vbNullPtr Then
        ReDim mBuffer(0 To mBufferSize - 1)
    End If
    
    If ReadFile(mHandle, mBuffer(0), mBufferSize, mBytesBuffered, ByVal 0&) = BOOL_FALSE Then
        If mHandle <> ERROR_BROKEN_PIPE Then
            IOError.WinIOError Err.LastDllError
        End If
    End If
End Sub

Private Function InternalSeek(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Currency
    With AsDLong(0.0001@ * Offset)
        .LoDWord = SetFilePointer(mHandle, .LoDWord, .HiDWord, Origin)
        ' in order to catch an error we check the return value against
        ' INVALID_SET_FILE_POINTER (&HFFFFFFFF). Even though it is a value
        ' of -1, the value must be treated as unsigned, so its negative
        ' values are = 2gig to 4gig positive values, which works with the
        ' 64bit positioning. So, if we get the -1, we need to check for
        ' an actual error using Err.LastDllError.
        If .LoDWord = INVALID_SET_FILE_POINTER Then
            Dim DllError As Long
            DllError = Err.LastDllError
            If DllError <> NO_ERROR Then
                IOError.WinIOError DllError
            End If
        End If
        InternalSeek = 10000@ * AsCurr(.LoDWord)
    End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    CloseStream
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
    Set Stream_BeginRead = BeginRead(Buffer, Offset, Count, Callback, State)
End Function

Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
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
    WriteByte Value
End Sub

Private Property Let Stream_WriteTimeout(ByVal Value As Long)
    WriteTimeout = Value
End Property

Private Property Get Stream_WriteTimeout() As Long
    Stream_WriteTimeout = WriteTimeout
End Property
