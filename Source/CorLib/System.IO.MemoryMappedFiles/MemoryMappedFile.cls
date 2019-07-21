VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedFile"
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
' Module: MemoryMappedFile
'

''
' Represents a simple file-to-memory mapping.
'
' @remarks Though this class is not part of .NET it is included
' to allow for easy access to mapping a file into memory.
' <p>A new <b>MemoryMappedFile</b> can be created using the <b>NewMemoryMappedFile</b>
' function.
' <pre>
'     Dim map As MemoryMappedFile
'     Set map = NewMemoryMappedFile("MyFile.txt")
' </pre>
' @see MemoryMappedFileStatic
Option Explicit
Implements IObject

Private mHandle     As SafeFileHandle
Private mStream     As FileStream
Private mLeaveOpen  As Boolean


''
' Gets the file handle of a memory-mapped file.
'
' @return The handle to the memory-mapped file.
Public Property Get SafeMemoryMappedFileHandle() As SafeFileHandle
    Set SafeMemoryMappedFileHandle = mHandle
End Property

''
' Creates a <a href="MemoryMappedViewAccessor.html">MemoryMappedViewAccessor</a> that maps to a view of the memory-mapped file, and that has the specified offset, size, and access restrictions.
'
' @param Offset The byte at which to start the view.
' @param Size The size of the view. Specify 0 (zero) to create a view that starts at <i>Offset</i> and ends approximately at the end of the memory-mapped file.
' @param Access One of the enumeration values that specifies the type of access allowed to the memory-mapped file. The default is <a href="MemoryMappedFileStatic.MemoryMappedFileAccess.html">ReadWrite</a>.
' @return A randomly accessible block of memory.
' @remarks <p>You can use the view returned by this method for random access to a memory-mapped file.</p>
' <p>To create a complete view of the memory-mapped file, specify 0 (zero) for the <i>Size</i> parameter. If you do this, the size of the view might be smaller or larger than the size of the source file on disk. This is because views are provided in units of system pages, and the size of the view is rounded up to the next system page size.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>Offset</i> or <i>Size</i> is a negative value.</td>
'   </tr>
' </table>
Public Function CreateViewAccessor(Optional ByVal Offset As Currency, Optional ByVal Size As Long, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedViewAccessor
    If Offset < 0 Then _
        Error.ArgumentOutOfRange "Offset", ArgumentOutOfRange_NeedNonNegNum
    If Size < 0 Then _
        Error.ArgumentOutOfRange "Size", ArgumentOutOfRange_PositiveOrDefaultSizeRequired
    If Access < MemoryMappedFileAccess.ReadWrite Or Access > MemoryMappedFileAccess.ReadWriteExecute Then _
        Error.ArgumentOutOfRange "Access"
    
    Set CreateViewAccessor = New MemoryMappedViewAccessor
    CreateViewAccessor.Init mHandle, Offset, Size, Access
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Handle As SafeFileHandle, ByVal Stream As FileStream, ByVal LeaveOpen As Boolean)
    Set mHandle = Handle
    Set mStream = Stream
    mLeaveOpen = LeaveOpen
End Sub

Private Sub Class_Terminate()
    If Not mLeaveOpen Then
        mStream.CloseStream
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
