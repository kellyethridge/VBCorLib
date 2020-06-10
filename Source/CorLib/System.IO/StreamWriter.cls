VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWriter"
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
' Module: StreamWriter
'

''
' Writes characters to a byte stream using the specific encoding.
'
' @remarks <b>StreamWriter</b> defaults to using an instance of <a href="UTF8Encoding.html">UTF8Encoding</a> unless specified otherwise.
' This instance of <a href="UTF8Encoding.html">UTF8Encoding</a> is constructed without a byte order mark (BOM), so its
' <a href="UTF8Encoding.GetPreamble.html">GetPreamble</a> method returns an empty byte array. To create a <b>StreamWriter</b>
' using UTF-8 encoding and a BOM, consider using a constructor that specifies encoding, such as <a href="Constructors.NewStreamWriter.html">NewStreamWriter(String, Boolean, Encoding)</a>.

' @see Constructors
' @see StreamReader
' @see TextWriter
' @include "..\DocIncludes\StreamWriter.html"
Option Explicit
Implements IObject
Implements TextWriter

Private Const MIN_BUFFERSIZE As Long = 128

Private mStream         As Stream
Private mAutoFlush      As Boolean
Private mEncoding       As Encoding
Private mEncoder        As Encoder
Private mNewLine        As String
Private mCharBuffer()   As Integer
Private mByteBuffer()   As Byte
Private mBufferSize     As Long
Private mPosition       As Long
Private mIsOpen         As Boolean
Private mPreambleWritten As Boolean
Private mLeaveOpen      As Boolean


''
' Returns if the writer autoflushes after each write.
'
' @return Indication of autoflush.
'
Public Property Get AutoFlush() As Boolean
    AutoFlush = mAutoFlush
End Property

''
' Sets if the writer autoflushes after each write.
'
' @param Value The new value.
'
Public Property Let AutoFlush(ByVal Value As Boolean)
    mAutoFlush = Value
    
    If mAutoFlush Then
        InternalFlush True, False
    End If
End Property

''
' Returns the stream this writer is using.
'
' @return The stream being used by this writer.
'
Public Property Get BaseStream() As Stream
    Set BaseStream = mStream
End Property

''
' Returns the encoding being used by this writer.
'
' @return The encoding being used by this writer.
'
Public Property Get Encoding() As Encoding
    Set Encoding = mEncoding
End Property

''
' Returns the format provider being used by this writer.
'
' @return The format provider used by this writer.
'
Public Property Get FormatProvider() As IFormatProvider
    Set FormatProvider = CultureInfo.CurrentCulture
End Property

''
' Returns the new-line string used by this writer.
'
' @return The new-line string.
'
Public Property Get NewLine() As String
    NewLine = mNewLine
End Property

''
' Sets the new-line string used by this writer.
'
' @param Value The new new-line value.
'
Public Property Let NewLine(ByVal Value As String)
    mNewLine = Value
End Property

''
' Closes this writer, flushing any buffers needed.
'
Public Sub CloseWriter()
    If mIsOpen Then
        Flush
        
        If Not mLeaveOpen Then
            On Error GoTo Finally
            mStream.CloseStream
        End If
    End If
    
Finally:
    Set mStream = Nothing
    Set mEncoder = Nothing
    Set mEncoding = Nothing
    Erase mCharBuffer
    Erase mByteBuffer
    mIsOpen = False
End Sub

''
' Flushes the buffers to the underlying stream.
'
Public Sub Flush()
    VerifyIsOpen
    InternalFlush True, True
End Sub

''
' Writes the encoded string representation of the value to the stream.
'
' @param value The value to be written to the stream.
' @param args The arguments to be used in formatting the Value.
' @remarks The <i>Args</i> are used when argument indexes are supplied
' in the value to be written. Indexes are indicated by enclosing the
' specific index in curly braces {}. The indexes start at 0.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The writer is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="FormatException.html">FormatException</a></td>
'        <td>
'           The format specification in format is invalid.<br/>
'           -or-<br/>
'           The number indicating an argument to be formatted is less than zero, or larger than or equal to the number of provided objects to be formatted.
'        </td>
'    </tr>
' </table>
Public Sub WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    VerifyIsOpen

    If UBound(Args) >= 0 Then
        Dim SwappedArgs() As Variant
        Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
        WriteFormatToStream Value, SwappedArgs
    Else
        WriteValueToStream Value
    End If
    
    If mAutoFlush Then
        InternalFlush True, False
    End If
End Sub

''
' Writes a subarray of characters to the stream.
'
' @param Buffer A character array containing the data to write.
' @param Index The character position in <i>Buffer</i> at which to start reading data.
' @param Count The maximum number of characters to write.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Buffer</i> is uninitialized.</td>
'    </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The writer is closed.</td>
'   </tr>
' </table>
Public Sub WriteChars(ByRef Buffer() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    Dim Output As String
    
    VerifyIsOpen
    
    If SAPtr(Buffer) = vbNullPtr Then _
        Error.ArgumentNull "Buffer", ArgumentNull_Array
    
    Output = Cor.NewString(Buffer, Index, Count)
    WriteCore Output
    
    If mAutoFlush Then
        InternalFlush True, False
    End If
End Sub

''
' Writes the value, then appends the new-line string.
'
' @param value The value to be written.
' @remarks The first argument in <i>Value</i> is the output to the
' string. Any additional arguments are used as formatting information.
' An argument can be refered to through formatting indexes. A
' formatting index is an integer enclosed in curly braces {}. The
' indexes start at 0.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The writer is closed.</td>
'   </tr>
'    <tr>
'        <td><a href="FormatException.html">FormatException</a></td>
'        <td>
'           The format specification in format is invalid.<br/>
'           -or-<br/>
'           The number indicating an argument to be formatted is less than zero, or larger than or equal to the number of provided objects to be formatted.
'        </td>
'    </tr>
' </table>
Public Sub WriteLine(ParamArray Value() As Variant)
    VerifyIsOpen
    
    If UBound(Value) >= 0 Then
        Dim SwappedArgs() As Variant
        Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(12)
        WriteFormatLine SwappedArgs
    Else
        WriteCore mNewLine
    End If
    
    If mAutoFlush Then
        InternalFlush True, False
    End If
End Sub

''
' Writes a subarray of characters to the text stream, followed by a line terminator.
'
' @param Buffer The character array from which data is read.
' @param Index The character position in <i>Buffer</i> at which to start reading data.
' @param Count The maximum number of characters to write.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td>The writer is closed.</td>
'   </tr>
' </table>
Public Sub WriteLineChars(ByRef Buffer() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    VerifyIsOpen
    
    WriteCore Cor.NewString(Buffer, Index, Count)
    WriteCore mNewLine
    
    If mAutoFlush Then
        InternalFlush True, False
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

' this is to allow Console to force the prevention
' of writting a preamble since it makes no sense.
Friend Sub HaveWrittenPreamble()
    mPreambleWritten = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Source As Variant, ByVal Append As Boolean, ByVal Encoding As Encoding, ByVal BufferSize As Long, ByVal LeaveOpen As Boolean)
    If BufferSize < 0 Then _
        Error.ArgumentOutOfRange "BufferSize", ArgumentOutOfRange_NeedNonNegNum
    
    Select Case VarType(Source)
        Case vbObject
            If Source Is Nothing Then _
                Error.ArgumentNull "Source", ArgumentNull_Stream
            If Not TypeOf Source Is Stream Then _
                Error.Argument Argument_StreamRequired, "Source"
            
            Set mStream = Source
            mLeaveOpen = LeaveOpen
            
        Case vbString
            Dim Mode As FileMode
            
            Mode = IIfLong(Append, FileMode.Append, FileMode.Create)
            Set mStream = Cor.NewFileStream(CStr(Source), Mode, FileAccess.WriteAccess, FileShare.ReadShare)
            
        Case Else
            Error.Argument Argument_InvalidStreamSource, "Source"
    End Select
    
    If mStream.CanSeek Then
        If mStream.Position > 0 Then
            HaveWrittenPreamble
        End If
    End If
    
    Set mEncoding = IfObject(Encoding, Stream.UTF8NoBOMAndThrows)
    Set mEncoder = mEncoding.GetEncoder
    mNewLine = Environment.NewLine
    mIsOpen = True
    mBufferSize = CorMath.Max(BufferSize, MIN_BUFFERSIZE)
    
    ReDim mCharBuffer(0 To mBufferSize - 1)
    ReDim mByteBuffer(0 To mEncoding.GetMaxByteCount(mBufferSize) - 1)
End Sub

Private Sub Class_Terminate()
    CloseWriter
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WriteFormatLine(ByRef Args() As Variant)
    Dim ArgsUBound As Long
    ArgsUBound = UBound(Args)
    
    Select Case ArgsUBound
        Case 0
            WriteValueToStream Args(0)
            
        Case Is > 0
            Dim Value As Variant
            Helper.MoveVariant Value, Args(0)
            CopyMemory Args(0), Args(1), (ArgsUBound + 1) * vbSizeOfVariant
            ZeroMemory Args(ArgsUBound), vbSizeOfVariant
            Args(ArgsUBound) = Empty
            WriteFormatToStream Value, Args
    
    End Select

    WriteCore mNewLine
End Sub

Private Sub WriteFormatToStream(ByRef Format As Variant, ByRef Args() As Variant)
    Dim Output As String
    Output = CorString.FormatArray(Convert.ToString(Format), Args)
    WriteCore Output
End Sub

Private Sub WriteValueToStream(ByRef Value As Variant)
    WriteCore Object.ToString(Value)
End Sub

Private Sub WriteCore(ByRef Value As String)
    Dim CharCount   As Long
    Dim CharPtr     As Long
    Dim CharsToCopy As Long
    
    CharCount = Len(Value)
    CharPtr = StrPtr(Value)
    
    Do While CharCount > 0
        ' If we will write past the end of the buffer, then
        ' empty the buffer in hopes we can fit inside it.
        If mPosition + CharCount >= mBufferSize Then
            InternalFlush False, False
        End If
        
        ' Determine just how many characters we can write to the buffer.
        If CharCount > mBufferSize Then
            ' Can't fit the entire thing in the buffer at once, so chop it up.
            CharsToCopy = mBufferSize
        Else
            ' It all fits now.
            CharsToCopy = CharCount
        End If
        
        ' Copy the characters to the buffer
        CopyMemory mCharBuffer(mPosition), ByVal CharPtr, CharsToCopy * 2

        ' Calculate the remaining characters to write.
        CharCount = CharCount - CharsToCopy
        CharPtr = CharPtr + CharsToCopy * 2
        mPosition = mPosition + CharsToCopy
    Loop
End Sub

Private Sub VerifyIsOpen()
    If Not mIsOpen Then
        Error.ObjectDisposed ObjectDisposed_StreamClosed
    End If
End Sub

Private Sub WritePreamble()
    Dim Bytes() As Byte
    
    If Not mPreambleWritten Then
        Bytes = mEncoding.GetPreamble
        
        If Len1D(Bytes) > 0 Then
            mStream.WriteBlock Bytes, LBound(Bytes), Len1D(Bytes)
        End If
        
        mPreambleWritten = True
    End If
End Sub

Private Sub InternalFlush(ByVal FlushStream As Boolean, ByVal FlushEncoder As Boolean)
    Dim Size As Long
    
    WritePreamble
    
    If mPosition > 0 Then
        Size = mEncoder.GetBytes(mCharBuffer, 0, mPosition, mByteBuffer, 0, FlushEncoder)
        mStream.WriteBlock mByteBuffer, 0, Size
        mPosition = 0
    End If
    
    If FlushStream Then
        mStream.Flush
    End If
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
'   TextWriter
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextWriter_CloseWriter()
    CloseWriter
End Sub

Private Property Get TextWriter_Encoding() As Encoding
    Set TextWriter_Encoding = Encoding
End Property

Private Function TextWriter_Equals(Value As Variant) As Boolean
    TextWriter_Equals = Equals(Value)
End Function

Private Sub TextWriter_Flush()
    Flush
End Sub

Private Property Get TextWriter_FormatProvider() As IFormatProvider
    Set TextWriter_FormatProvider = FormatProvider
End Property

Private Function TextWriter_GetHashCode() As Long
    TextWriter_GetHashCode = GetHashCode
End Function

Private Property Let TextWriter_NewLine(ByVal Value As String)
    NewLine = Value
End Property

Private Property Get TextWriter_NewLine() As String
    TextWriter_NewLine = NewLine
End Property

Private Function TextWriter_ToString() As String
    TextWriter_ToString = ToString
End Function

Private Sub TextWriter_WriteChars(Chars() As Integer, Optional Index As Variant, Optional Count As Variant)
    WriteChars Chars, Index, Count
End Sub

Private Sub TextWriter_WriteLine(ParamArray Value() As Variant)
    VerifyIsOpen
    
    If UBound(Value) >= 0 Then
        Dim SwappedArgs() As Variant
        Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(12)
        WriteFormatLine SwappedArgs
    Else
        WriteCore mNewLine
    End If
    
    If mAutoFlush Then
        InternalFlush True, False
    End If
End Sub

Private Sub TextWriter_WriteLineChars(Chars() As Integer, Optional Index As Variant, Optional Count As Variant)
    WriteLineChars Chars, Index, Count
End Sub

Private Sub TextWriter_WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    VerifyIsOpen

    If UBound(Args) >= 0 Then
        Dim SwappedArgs() As Variant
        Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
        WriteFormatToStream Value, SwappedArgs
    Else
        WriteValueToStream Value
    End If
    
    If mAutoFlush Then
        InternalFlush True, False
    End If
End Sub

