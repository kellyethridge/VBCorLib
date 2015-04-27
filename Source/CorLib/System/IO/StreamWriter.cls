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
' @remarks The default encoding used is UTF8Encoding. If the writer is
' not appending to the stream, then an encoding preamble is written to
' the stream.
'
' @see Constructors
' @see StreamReader
' @see TextWriter
'
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
' @param RHS The new value.
'
Public Property Let AutoFlush(ByVal Value As Boolean)
    mAutoFlush = Value
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
' @param RHS The new new-line value.
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
        mStream.CloseStream
        mIsOpen = False
    End If
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
' @remarks The <i>args</i> are used when argument indexes are supplied
' in the value to be written. Indexes are indicate by enclosing the
' specific index in curly braces {}. The indexes start at 0.
'
Public Sub WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    InternalWriteValue Value, SwappedArgs
End Sub

Public Sub WriteChars(ByRef Chars() As Integer)
    If SAPtr(Chars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
        
    Dim Output As String
    Output = cString.FromCharArray(Chars)
    InternalWrite Output
End Sub


''
' Writes the value, then appends the new-line string.
'
' @param value The value to be written.
' @remarks The first argument in <i>value</i> is the output to the
' string. Any additional arguments are used as formatting information.
' An argument can be refered to through formatting indexes. A
' formatting index is an integer enclosed in curly braces {}. The
' indexes start at 0.
'
Public Sub WriteLine(ParamArray Value() As Variant)
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(12)
    InternalWriteLine SwappedArgs
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
Friend Sub Init(ByVal Source As Variant, ByVal Append As Boolean, ByVal Encoding As Encoding, ByVal BufferSize As Long)
    Dim Mode As FileMode
    Dim EmitBOM As Boolean
    
    If BufferSize < 0 Then _
        Error.NegativeNumber "BufferSize"
    
    Select Case VarType(Source)
        Case vbObject
            If Source Is Nothing Then _
                Error.ArgumentNull "Source", ArgumentNull_Stream
            If Not TypeOf Source Is Stream Then _
                Error.Argument Argument_StreamRequired, "Source"
                
            Set mStream = Source
        
        Case vbString
            If Append Then
                Mode = FileMode.Append
            Else
                Mode = FileMode.Create
                EmitBOM = True
            End If
            
            Set mStream = Cor.NewFileStream(Source, Mode, FileAccess.WriteAccess, FileShare.ReadShare)
        
        Case Else
            Error.Argument Argument_InvalidStreamSource, "Source"
    End Select
        
    If Encoding Is Nothing Then
        Set mEncoding = Cor.NewUTF8Encoding(False, True)
    Else
        Set mEncoding = Encoding
    End If
    Set mEncoder = mEncoding.GetEncoder
    
    mNewLine = Environment.NewLine
    mIsOpen = True
    
    If BufferSize < MIN_BUFFERSIZE Then
        mBufferSize = MIN_BUFFERSIZE
    Else
        mBufferSize = BufferSize
    End If
    
    ReDim mCharBuffer(0 To mBufferSize - 1)
    ReDim mByteBuffer(0 To mEncoding.GetMaxByteCount(mBufferSize) - 1)
    
    If EmitBOM Then
        WritePreamble
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InternalWrite(ByVal Value As String, ByRef Args() As Variant, ByVal NewLine As Boolean)
    Dim CharCount As Long
    Dim CharPtr As Long
    Dim CharsToCopy As Long
    
    Call VerifyIsOpen
    
    If UBound(Args) >= 0 Then Value = cString.FormatArray(Value, Args)
    If NewLine Then Value = Value & mNewLine
    
    CharCount = Len(Value)
    CharPtr = StrPtr(Value)
    
    Do While CharCount > 0
        ' If we will write past the end of the buffer, then
        ' empty the buffer in hopes we can fit inside it.
        If mPosition + CharCount >= mBufferSize Then Call InternalFlush(False, False)
        
        ' Determine just how many characters we can write to the buffer.
        If CharCount > mBufferSize Then
            ' Can't fit the entire thing in the buffer at once, so chop it up.
            CharsToCopy = mBufferSize
        Else
            ' It all fits now.
            CharsToCopy = CharCount
        End If
        
        ' Copy the characters to the buffer
        Call CopyMemory(mCharBuffer(mPosition), ByVal CharPtr, CharsToCopy * 2)
        
        ' Calculate the remaining characters to write.
        CharCount = CharCount - CharsToCopy
        CharPtr = CharPtr + CharsToCopy * 2
        mPosition = mPosition + CharsToCopy
    Loop
    If mAutoFlush Then Call InternalFlush(True, False)
End Sub

' WriteLine has only a ParamArray argument. This allows for nothing
' to be passed in and only write a newline to the stream.
Private Sub InternalWriteLine(ByRef Args() As Variant)
    Dim ub As Long
    Dim Value As String
    
    ub = UBound(Args)
    
    ' The first argument is the line to be written, so convert it to a string.
    If ub >= 0 Then Value = Convert.ToString(Args(0))
    
    ' If there are additional arguments, then those arguments need
    ' to be shifted down to fill the space of the first argument, which
    ' is now a separate value.
    If ub > 0 Then
        Dim i As Long
        For i = 1 To ub
            Call Helper.MoveVariant(Args(i - 1), Args(i))
        Next i
        ReDim Preserve Args(ub - 1)
    ElseIf ub = 0 Then
        ' The first element is now the Value and we just
        ' want an empty zero-length array now.
        Args = cArray.CreateInstance(vbVariant)
    End If
    
    Call InternalWrite(Value, Args, True)
End Sub

Private Sub InternalWriteValue(ByRef Value As Variant, ByRef Args() As Variant)
    Call InternalWrite(Convert.ToString(Value), Args, False)
End Sub

Private Sub VerifyIsOpen()
    If Not mIsOpen Then Throw Cor.NewObjectDisposedException("StreamWriter", "Cannot write to a close writer.")
End Sub

Private Sub WritePreamble()
    Dim b() As Byte
    
    If Not mStream.CanSeek Then Exit Sub
    If mStream.Length > 0 Then Exit Sub
    
    b = mEncoding.GetPreamble
    Call mStream.WriteBlock(b, 0, cArray.GetLength(b))
End Sub

Private Sub InternalFlush(ByVal FlushStream As Boolean, ByVal FlushEncoder As Boolean)
    Dim Size As Long
    
    If mPosition > 0 Then
        Size = mEncoder.GetBytes(mCharBuffer, 0, mPosition, mByteBuffer, 0, FlushEncoder)
        Call mStream.WriteBlock(mByteBuffer, 0, Size)
        mPosition = 0
    End If
    
    If FlushStream Then Call mStream.Flush
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call CloseWriter
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
'   TextWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextWriter_CloseWriter()
    Call CloseWriter
End Sub

Private Property Get TextWriter_Encoding() As Encoding
    Set TextWriter_Encoding = Encoding
End Property

Private Function TextWriter_Equals(Value As Variant) As Boolean
    TextWriter_Equals = Equals(Value)
End Function

Private Sub TextWriter_Flush()
    Call Flush
End Sub

Private Property Get TextWriter_FormatProvider() As IFormatProvider
    Set TextWriter_FormatProvider = FormatProvider
End Property

Private Function TextWriter_GetHashCode() As Long
    TextWriter_GetHashCode = GetHashCode
End Function

Private Property Let TextWriter_NewLine(ByVal RHS As String)
    NewLine = RHS
End Property

Private Property Get TextWriter_NewLine() As String
    TextWriter_NewLine = NewLine
End Property

Private Function TextWriter_ToString() As String
    TextWriter_ToString = ToString
End Function

Private Sub TextWriter_WriteLine(ParamArray Value() As Variant)
    Dim a() As Variant
    Call Helper.Swap4(ByVal ArrPtr(a), ByVal Helper.DerefEBP(12))
    Call InternalWriteLine(a)
End Sub

Private Sub TextWriter_WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    Dim a() As Variant
    Call Helper.Swap4(ByVal ArrPtr(a), ByVal Helper.DerefEBP(16))
    Call InternalWriteValue(Value, a)
End Sub

