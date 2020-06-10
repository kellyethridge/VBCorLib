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
' @remarks A memory-mapped file maps the contents of a file to an application’s logical address space. Memory-mapped files enable programmers to
' work with extremely large files because memory can be managed concurrently, and they allow complete, random access to a file without the need
' for seeking. Memory-mapped files can also be shared across multiple processes.
' <p>The <a href="MemoryMappedFileStatic.CreateFromFile.html">CreateFromFile</a> and <a href="MemoryMappedFileStatic.CreateFromFileStream.html">CreateFromFileStream</a>
' methods create a memory-mapped file from a specified path or a <a href="FileStream.html">FileStream</a> of an existing file on disk. Changes are automatically propagated to disk when the file is unmapped.</p>
' <p>The <a href="MemoryMappedFileStatic.CreateNew.html">CreateNew</a> method creates a memory-mapped file that is not mapped to an existing file on disk; and are suitable for creating shared memory for interprocess communication (IPC).</p>
' <p>A memory-mapped file is associated with a name.</p>
' <p>You can create multiple views of the memory-mapped file, including views of parts of the file. You can map the same part of a file to more than one
' address to create concurrent memory. For two views to remain concurrent, they have to be created from the same memory-mapped file. Creating two file
' mappings of the same file with two views does not provide concurrency.</p>
' @see MemoryMappedFileStatic
' @include "..\DocIncludes\MemoryMappedFile.html"
Option Explicit
Implements IObject

Private mHandle     As SafeMemoryMappedFileHandle
Private mStream     As FileStream
Private mLeaveOpen  As Boolean

''
' Specifies access capabilities and restrictions for a memory-mapped file or view.
'
' @param CopyOnWrite Read and write access to the file, with the restriction that any write operations will not be seen by other processes.
' @param ReadOnly Read-only access to the file.
' @param ReadExecute Read access to the file that can store and run executable code.
' @param ReadWrite Read and write access to the file.
' @param ReadWriteExecute Read and write access to the file that can store and run executable code.
' @param WriteOnly Write-only access to file.
Public Enum MemoryMappedFileAccess
    ReadWrite = 0
    ReadOnly
    WriteOnly   ' WriteOnly is valid only when creating views and not when creating MemoryMappedFiles
    CopyOnWrite
    ReadExecute
    ReadWriteExecute
End Enum

''
' Specifies access rights to a memory-mapped file that is not associated with a file on disk.
'
' @param CopyOnWrite The right to read and write to a file with the restriction that write operations will not be seen by other processes.
' @param [Write] The right to add data to a file or remove data from a file.
' @param Read The right to open and copy a file as read-only.
' @param Execute The right to run an application file.
' @param Delete The right to delete a file.
' @param ReadPermissions The right to open and copy access and audit rules from a file. This does not include the right to read data, file system attributes, or extended file system attributes.
' @param ChangePermissions The right to change the security and audit rules associated with a file.
' @param TakeOwnership The right to change the owner of a file.
' @param ReadWrite The right to open and copy a file, and the right to add data to a file or remove data from a file.
' @param ReadExecute The right to open and copy a folder or file as read-only, and to run application files. This right includes the <b>Read</b> right and the <b>Execute</b> right.
' @param ReadWriteExecute The right to open and copy a file, the right to add data to a file or remove data from a file, and the right to run an application file.
' @param FullControl The right to exert full control over a file, and to modify access control and audit rules. This value represents the right to do anything with a file and is the combination of all rights in this enumeration.
' @param AccessSystemSecurity The right to get or set permissions on a file.
Public Enum MemoryMappedFileRights
    CopyOnWrite = &H1
    [Write] = &H2
    Read = &H4
    Execute = &H8
    Delete = &H10000
    ReadPermissions = &H20000
    ChangePermissions = &H40000
    TakeOwnership = &H80000
    ReadWrite = MemoryMappedFileRights.Read Or MemoryMappedFileRights.Write
    ReadExecute = MemoryMappedFileRights.Read Or MemoryMappedFileRights.Execute
    ReadWriteExecute = MemoryMappedFileRights.Read Or MemoryMappedFileRights.Write Or MemoryMappedFileRights.Execute
    FullControl = MemoryMappedFileRights.CopyOnWrite Or MemoryMappedFileRights.Read Or MemoryMappedFileRights.Write Or MemoryMappedFileRights.Execute Or MemoryMappedFileRights.Delete Or MemoryMappedFileRights.ReadPermissions Or MemoryMappedFileRights.ChangePermissions Or MemoryMappedFileRights.TakeOwnership
    AccessSystemSecurity = &H1000000
End Enum


''
' Gets the file handle of a memory-mapped file.
'
' @return The handle to the memory-mapped file.
Public Property Get SafeMemoryMappedFileHandle() As SafeMemoryMappedFileHandle
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
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'       <td>
'           <i>Offset</i> or <i>Size</i> is a negative value.<br/>
'           -or-<br/>
'           <i>Access</i> is not a valid <a href="MemoryMappedFile.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> enumeration value.
'       </td>
'   </tr>
'    <tr>
'       <td><a href="UnauthorizedAccessException.html">UnauthorizedAccessException</a></td><td><i>Access</i> is invalid for the memory-mapped file.</td>
'    </tr>
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
' Creates a stream that maps to a view of the memory-mapped file, and that has the specified offset, size, and access type.
'
' @param Offset The byte at which to start the view.
' @param Size The size of the view. Specify 0 (zero) to create a view that starts at offset and ends approximately at the end of the memory-mapped file.
' @param Access One of the enumeration values that specifies the type of access allowed to the memory-mapped file. The default is <a href="MemoryMappedFile.MemoryMappedFileAccess.html">ReadWrite</a>.
' @return A stream of memory that has the specified characteristics.
' @remarks <p>You can use the stream returned by this method for sequential access to a memory-mapped file, such as for inter-process communications.</p>
' <p>To create a complete view of the memory-mapped file, specify 0 (zero) for the <i>Size</i> parameter. If you do this, the size of the view might be smaller or larger than the size of the source file on disk. This is because views are provided in units of system pages, and the size of the view is rounded up to the next system page size.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>
'           <i>Offset</i> or <i>Size</i> is a negative value.<br/>
'           -or-<br/>
'           <i>Access</i> is not a valid <a href="MemoryMappedFile.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> enumeration value.
'       </td>
'   </tr>
'    <tr>
'       <td><a href="UnauthorizedAccessException.html">UnauthorizedAccessException</a></td><td><i>Access</i> is invalid for the memory-mapped file.</td>
'    </tr>
' </table>
Public Function CreateViewStream(Optional ByVal Offset As Currency, Optional ByVal Size As Long, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedViewStream
    If Offset < 0 Then _
        Error.ArgumentOutOfRange "Offset", ArgumentOutOfRange_NeedNonNegNum
    If Size < 0 Then _
        Error.ArgumentOutOfRange "Size", ArgumentOutOfRange_PositiveOrDefaultSizeRequired
    If Access < MemoryMappedFileAccess.ReadWrite Or Access > MemoryMappedFileAccess.ReadWriteExecute Then _
        Error.ArgumentOutOfRange "Access"
    
    Set CreateViewStream = New MemoryMappedViewStream
    CreateViewStream.Init mHandle, Offset, Size, Access
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
Friend Sub Init(ByVal Handle As SafeMemoryMappedFileHandle, ByVal Stream As FileStream, ByVal LeaveOpen As Boolean)
    Set mHandle = Handle
    Set mStream = Stream
    mLeaveOpen = LeaveOpen
End Sub

Private Sub Class_Terminate()
    If Not mLeaveOpen Then
        If Not mStream Is Nothing Then
            mStream.CloseStream
        End If
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
