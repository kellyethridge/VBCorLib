VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedFileStatic"
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
' Module: MemoryMappedFileStatic
'

''
' Creates a memory-mapped file that has the specified access mode, name, capacity, and access type from a file on disk.
' @see MemoryMappedFile
Option Explicit

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
' Creates a memory-mapped file that has the specified access mode, name, capacity, and access type from a file on disk.
'
' @param Path The path to the file to map.
' @param Mode Access mode; can be any of the <a href="FileStream.FileMode.html">FileMode</a> enumeration values except <a href="FileStream.FileMode.html">Append</a>.
' @param MapName A name to assign to the memory-mapped file.
' @param Capacity The maximum size, in bytes, to allocate to the memory-mapped file. Specify 0 to set the capacity to the size of the file on disk.
' @param Access One of the enumeration values that specifies the type of access allowed to the memory-mapped file.
' @return A memory-mapped file that has the specified characteristics.
' @remarks <p>The <i>mode</i> parameter pertains to the source file on disk.</p>
' <p>If <i>Capacity</i> is larger than the size of the file on disk, the file on disk is increased to match the specified capacity even if no data is written to the memory-mapped file. To prevent this from occurring, specify 0 (zero) for the default capacity, which will internally set capacity to the size of the file on disk.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td>
'       <td>
'           <i>MapName</i> is an empty string.<br/>
'           -or-
'           <i>Access</i> is not an allowed value.<br/>
'           -or-<br/>
'           <i>Path</i> specifies an empty file.<br/>
'           -or-<br/>
'           <i>Access</i> is specified as <a href="MemoryMappedFileStatic.MemoryMappedFileAccess.html">ReadOnly</a> and <i>Capacity</i> is greater than the size of the file indicated by <i>Path</i>.<br/>
'           -or-<br/>
'           <i>Mode</i> is <a href="FileStream.FileMode.html">Append</a>.
'       </td>
'   </tr>
'    <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>
'           <i>Capacity</i> is less than zero.<br/>
'           -or-<br/>
'           <i>Capacity</i> is less than the file size (but not zero).<br/>
'           -or-<br/>
'           <i>Capacity</i> is zero, and the size of the file on disk is also zero.<br/>
'           -or-<br/>
'           <i>Access</i> is not a defined <a href="MemoryMappedFileStatic.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> value.<br/>
'           -or-<br/>
'           The size of the file indicated by <i>Path</i> is greater than <i>Capacity</i>.
'       </td>
'    </tr>
' </table>
Public Function CreateFromFile(ByRef Path As String, Optional ByVal Mode As FileMode = FileMode.OpenExisting, Optional ByRef MapName As String, Optional ByVal Capacity As Currency, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    Dim Existed As Boolean
    Dim Stream  As FileStream
    Dim Handle  As SafeFileHandle
    
    If Not CorString.IsNull(MapName) And Len(MapName) = 0 Then _
        Error.Argument Argument_MapNameEmptyString
    If Capacity < 0 Then _
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_PositiveOrDefaultCapacityRequired
    If Access < MemoryMappedFileAccess.ReadWrite Or Access > MemoryMappedFileAccess.ReadWriteExecute Then _
        Error.ArgumentOutOfRange "Access"
    If Mode = FileMode.Append Then _
        Error.Argument Argument_NewMMFAppendModeNotAllowed, "Mode"
    If Access = MemoryMappedFileAccess.WriteOnly Then _
        Error.Argument Argument_NewMMFWriteAccessNotAllowed, "Access"
    
    Existed = File.Exists(Path)
    Set Stream = Cor.NewFileStream(Path, Mode, GetFileAccess(Access), FileShare.None)
    
    If Capacity = 0 And Stream.Length = 0 Then
        CleanupFile Stream, Existed, Path
        Error.Argument Argument_EmptyFile
    End If
    
    If Capacity = 0 Then
        Capacity = Stream.Length
    ElseIf Stream.Length > Capacity Then
        CleanupFile Stream, Existed, Path
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_CapacityGEFileSizeRequired
    ElseIf Capacity > Stream.Length And Access = MemoryMappedFileAccess.ReadOnly Then
        CleanupFile Stream, Existed, Path
        Error.Argument Argument_ReadAccessWithLargeCapacity
    End If

    On Error GoTo Catch
    Set Handle = CreateCore(Stream.SafeFileHandle, MapName, Access, Capacity)
    Set CreateFromFile = Cor.NewMemoryMappedFile(Handle, Stream, False)
    Exit Function

Catch:
    CleanupFile Stream, Existed, Path
    ThrowOrErr Err
End Function

Private Function CreateCore(ByVal FileHandle As SafeFileHandle, ByRef MapName As String, ByVal Access As MemoryMappedFileAccess, ByVal Capacity As Currency) As SafeFileHandle
    Dim Handle          As SafeFileHandle
    Dim CapacityHiLow   As DLong
    Dim ErrCode         As Long
    
    Capacity = Capacity / 10000@ ' move all bits to start at zero bit.
    CapacityHiLow = AsDLong(Capacity)
    Set Handle = SafeCreateFileMapping(FileHandle.DangerousGetHandle, GetPageAccess(Access), CapacityHiLow.HiDWord, CapacityHiLow.LoDWord, MapName)
    ErrCode = Err.LastDllError
    
    If Not Handle.IsInvalid And ErrCode = ERROR_ALREADY_EXISTS Then
        Handle.CloseHandle
        IOError.WinIOError ErrCode
    ElseIf Handle.IsInvalid Then
        IOError.WinIOError ErrCode
    End If
    
    Set CreateCore = Handle
End Function

Private Function GetFileAccess(ByVal Access As MemoryMappedFileAccess) As FileAccess
    Select Case Access
        Case MemoryMappedFileAccess.ReadOnly, MemoryMappedFileAccess.ReadExecute
            GetFileAccess = ReadAccess
        Case MemoryMappedFileAccess.WriteOnly
            GetFileAccess = WriteAccess
        Case MemoryMappedFileAccess.ReadWrite, MemoryMappedFileAccess.CopyOnWrite, MemoryMappedFileAccess.ReadWriteExecute
            GetFileAccess = ReadWriteAccess
        Case Else
            Error.ArgumentOutOfRange "Access"
    End Select
End Function

Private Function GetPageAccess(ByVal Access As MemoryMappedFileAccess) As Long
    Select Case Access
        Case MemoryMappedFileAccess.ReadOnly
            GetPageAccess = PAGE_READONLY
        Case MemoryMappedFileAccess.ReadWrite
            GetPageAccess = PAGE_READWRITE
        Case MemoryMappedFileAccess.CopyOnWrite
            GetPageAccess = PAGE_WRITECOPY
        Case MemoryMappedFileAccess.ReadExecute
            GetPageAccess = PAGE_EXECUTE_READ
        Case MemoryMappedFileAccess.ReadWriteExecute
            GetPageAccess = PAGE_EXECUTE_READWRITE
        Case Else
            Error.ArgumentOutOfRange "Access"
    End Select
End Function

Private Sub CleanupFile(ByVal Stream As FileStream, ByVal Existed As Boolean, ByRef Path As String)
    Stream.CloseStream
    
    If Not Existed Then
        File.Delete Path
    End If
End Sub

