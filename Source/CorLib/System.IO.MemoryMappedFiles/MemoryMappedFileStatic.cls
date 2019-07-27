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
' Creates a memory-mapped file that has the specified access mode, name, capacity, and access type from a file on disk.
'
' @param Path The path to the file to map.
' @param Mode Access mode; can be any of the <a href="FileStream.FileMode.html">FileMode</a> enumeration values except <a href="FileStream.FileMode.html">Append</a>.
' @param MapName A name to assign to the memory-mapped file. Specify <b>vbNullString</b> to rename unnamed.
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
'           -or-<br/>
'           <i>Access</i> is not an allowed value.<br/>
'           -or-<br/>
'           <i>Path</i> specifies an empty file.<br/>
'           -or-<br/>
'           <i>Access</i> is specified as <a href="MemoryMappedFile.MemoryMappedFileAccess.html">ReadOnly</a> and <i>Capacity</i> is greater than the size of the file indicated by <i>Path</i>.<br/>
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
'           <i>Access</i> is not a defined <a href="MemoryMappedFile.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> value.<br/>
'           -or-<br/>
'           The size of the file indicated by <i>Path</i> is greater than <i>Capacity</i>.
'       </td>
'    </tr>
' </table>
Public Function CreateFromFile(ByRef Path As String, Optional ByVal Mode As FileMode = FileMode.OpenExisting, Optional ByRef MapName As String, Optional ByVal Capacity As Currency, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    Dim Existed As Boolean
    Dim Stream  As FileStream
    Dim Handle  As SafeMemoryMappedFileHandle
    
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
    
    Capacity = Int(Capacity)
    Existed = File.Exists(Path)
    Set Stream = Cor.NewFileStream(Path, Mode, GetFileAccess(Access), FileShare.None)
    
    If Capacity = 0 And Stream.Length = 0 Then
        CleanupFile Stream, Existed, Path
        Error.Argument Argument_EmptyFile
    End If
    
    If Capacity = 0 Then
        Capacity = Stream.Length
    ElseIf Capacity < Stream.Length Then
        CleanupFile Stream, Existed, Path
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_CapacityGEFileSizeRequired
    ElseIf Capacity > Stream.Length And Access = MemoryMappedFileAccess.ReadOnly Then
        CleanupFile Stream, Existed, Path
        Error.Argument Argument_ReadAccessWithLargeCapacity
    End If

    On Error GoTo Catch
    Set Handle = CreateCore(Stream.SafeFileHandle, MapName, Capacity, Access)
    Set CreateFromFile = Cor.NewMemoryMappedFile(Handle, Stream, False)
    Exit Function

Catch:
    CleanupFile Stream, Existed, Path
    ThrowOrErr Err
End Function

''
' Creates a memory-mapped file that has the specified name, capacity, access type, and disposal requirement from a file on disk.
'
' @param FileStream The <i>FileStream</i> to the file to map.
' @param MapName A name to assign to the memory-mapped file. Specify <b>vbNullString</b> to leave mapping unnamed.
' @param Capacity The maximum size, in bytes, to allocate to the memory-mapped file. Specify 0 to set the capacity to the size of the file on disk.
' @param Access One of the enumeration values that specifies the type of access allowed to the memory-mapped file. <p>This parameter can't be set to <a href="MemoryMappedFile.MemoryMappedFileAccess.html">WriteOnly</a>.</p>
' @param LeaveOpen <b>True</b> to not dispose fileStream after the <a href="MemoryMappedFile.html">MemoryMappedFile</a> is closed; <b>False</b> to dispose <i>FileStream</i>.
' @return A memory-mapped file that has the specified characteristics.
' @remarks If <i>Capacity</i> is larger than the size of the file on disk, the file on disk is increased to match the specified capacity even if no data is written to the memory-mapped file. To prevent this from occurring, specify 0 (zero) for the default capacity, which will internally set <i>Capacity</i> to the size of the file on disk.
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
'           -or-<br/>
'           <i>Capacity</i> and the length of the file are zero.<br/>
'           -or-<br/>
'           <i>Access</i> is set to the <a href="MemoryMappedFile.MemoryMappedFileAccess.html">WriteOnly</a> enumeration value, which is not allowed.<br/>
'           -or-<br/>
'           <i>Access</> is set to the <a href="MemoryMappedFile.MemoryMappedFileAccess.html">ReadOnly</a> enumeration value and <i>Capacity</i> is larger than file size.
'       </td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>FileStream</i> is <b>Nothing</b>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>Capacity</i> is less than zero.<br/>
'           -or-<br/>
'           <i>Capacity</i> is less than the file size.<br/>
'           -or-<br/>
'           <i>Access</i> is not a valid <a href="MemoryMappedFile.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> enumeration value.<br/>
'        </td>
'    </tr>
'    <tr>
'        <td><a href="ObjectDisposedException.html">ObjectDisposedException</a></td><td><i>FileStream</i> was closed.</td>
'    </tr>
'    <tr>
'        <td><a href="IOException.html">IOException</a></td><td><i>MapName</i> already exists.</td>
'    </tr>
' </table>
Public Function CreateFromFileStream(ByVal FileStream As FileStream, ByRef MapName As String, ByVal Capacity As Currency, ByVal Access As MemoryMappedFileAccess, ByVal LeaveOpen As Boolean) As MemoryMappedFile
    Dim Handle As SafeMemoryMappedFileHandle
    
    If FileStream Is Nothing Then _
        Error.ArgumentNull "FileStream"
    If Not CorString.IsNull(MapName) And Len(MapName) = 0 Then _
        Error.Argument Argument_MapNameEmptyString
    If Capacity < 0 Then _
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_PositiveOrDefaultCapacityRequired
    If Access < MemoryMappedFileAccess.ReadWrite Or Access > MemoryMappedFileAccess.ReadWriteExecute Then _
        Error.ArgumentOutOfRange "Access"
    
    Capacity = Int(Capacity)
    
    If Capacity = 0 And FileStream.Length = 0 Then _
        Error.Argument Argument_EmptyFile
    If Access = MemoryMappedFileAccess.WriteOnly Then _
        Error.Argument Argument_NewMMFWriteAccessNotAllowed, "Access"
    
    FileStream.Flush
    
    If Capacity = 0 Then
        Capacity = FileStream.Length
    End If
    
    If Capacity < FileStream.Length Then _
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_CapacityGEFileSizeRequired
    If Capacity > FileStream.Length And Access = MemoryMappedFileAccess.ReadOnly Then _
        Error.Argument Argument_ReadAccessWithLargeCapacity
    
    Set Handle = CreateCore(FileStream.SafeFileHandle, MapName, Capacity, Access)
    Set CreateFromFileStream = Cor.NewMemoryMappedFile(Handle, FileStream, LeaveOpen)
End Function

''
' Creates a memory-mapped file that has the specified capacity and access type in system memory.
'
' @param MapName A name to assign to the memory-mapped file.
' @param Capacity The maximum size, in bytes, to allocate to the memory-mapped file.
' @param Access One of the enumeration values that specifies the type of access allowed to the memory-mapped file. The default is <a href="MemoryMappedFile.MemoryMappedFileAccess.html">ReadWrite</a>.
' @return A memory-mapped file that has the specified characteristics.
' @remarks Use this method to create a memory-mapped file that is not persisted (that is, not associated with a file on disk), which you can use to share data between processes.
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
'           -or-<br/>
'           <i>Access</i> is set to write-only with the <a href="MemoryMappedFile.MemoryMappedFileAccess.html">WriteOnly</a> enumeration value.
'       </td>
'   </tr>
'    <tr>
'       <td></td>
'       <td>
'           <i>Capacity</i> is less than or equal to zero.<br/>
'           -or-<br/>
'           <i>Access</i> is not a valid <a href="MemoryMappedFile.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> enumeration value.
'       </td>
'    </tr>
' </table>
Public Function CreateNew(ByRef MapName As String, ByVal Capacity As Currency, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    Dim Handle As SafeMemoryMappedFileHandle
    
    If Not CorString.IsNull(MapName) And Len(MapName) = 0 Then _
        Error.Argument Argument_MapNameEmptyString
    
    Capacity = Int(Capacity)
    
    If Capacity <= 0 Then _
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_NeedPosNum
    If Access < MemoryMappedFileAccess.ReadWrite Or Access > MemoryMappedFileAccess.ReadWriteExecute Then _
        Error.ArgumentOutOfRange "Access"
    If Access = MemoryMappedFileAccess.WriteOnly Then _
        Error.Argument Argument_NewMMFWriteAccessNotAllowed, "Access"

    Set Handle = CreateCore(Cor.NewSafeFileHandle(-1, True), MapName, Capacity, Access)
    Set CreateNew = Cor.NewMemoryMappedFile(Handle)
End Function

''
' Opens an existing memory-mapped file that has the specified name in system memory.
'
' @param MapName The name of the memory-mapped file to open.
' @param DesiredAccessRights One of the enumeration values that specifies the access rights to apply to the memory-mapped file.
' @return A memory-mapped file that has the specified name.
' @remarks The memory-mapped file can be either a persisted memory-mapped file (associated with a file on disk) or non persisted.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>MapName</i> is an empty string.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>MapName</i> is <b>vbNullString</b>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>DesiredAccessRights</i> is not a valid <a href="MemoryMappedFile.MemoryMappedFileRights.html">MemoryMappedFileRights</a> enumeration value.</td>
'    </tr>
'    <tr>
'        <td><a href="FileNotFoundException.html">FileNotFoundException</td><td>The file specified for <i>MapName</i> does not exist.</td>
'    </tr>
' </table>
Public Function OpenExisting(ByRef MapName As String, Optional ByVal DesiredAccessRights As MemoryMappedFileRights = MemoryMappedFileRights.ReadWrite) As MemoryMappedFile
    Dim Handle As SafeMemoryMappedFileHandle
    
    If CorString.IsNull(MapName) Then _
        Error.ArgumentNull "MapName"
    If Len(MapName) = 0 Then _
        Error.Argument Argument_MapNameEmptyString
    If (DesiredAccessRights And Not (MemoryMappedFileRights.FullControl Or MemoryMappedFileRights.AccessSystemSecurity)) <> 0 Then _
        Error.ArgumentOutOfRange "DesiredAccessRights"
    
    Set Handle = OpenCore(MapName, DesiredAccessRights, False)
    Set OpenExisting = Cor.NewMemoryMappedFile(Handle)
End Function

''
' Creates or opens a memory-mapped file that has the specified capacity and access type in system memory.
'
' @param MapName A name to assign to the memory-mapped file.
' @param Capacity The maximum size, in bytes, to allocate to the memory-mapped file.
' @param Access One of the enumeration values that specifies the type of access allowed to the memory-mapped file. The default is <a href="MemoryMappedFile.MemoryMappedFileAccess.html">ReadWrite</a>.
' @return A memory-mapped file that has the specified characteristics.
' @remarks Use this method to create or open a memory-mapped file that is not persisted (that is, not associated with a file on disk), which you can use to share data between processes.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>MapName</i> is an empty string.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>MapName is <b>vbNullString</b>.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>Capacity</i> is greater than the size of the logical address space.<br/>
'           -or-<br/>
'           <i>Capacity</i> is less than or equal to zero.<br/>
'           -or-<br/>
'           <i>Access</i> is not a valid <a href="MemoryMappedFile.MemoryMappedFileAccess.html">MemoryMappedFileAccess</a> enumeration value.
'        </td>
'    </tr>
' </table>
Public Function CreateOrOpen(ByRef MapName As String, ByVal Capacity As Currency, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    Dim Handle As SafeMemoryMappedFileHandle
    
    If CorString.IsNull(MapName) Then _
        Error.ArgumentNull "MapName"
    If Len(MapName) = 0 Then _
        Error.Argument Argument_MapNameEmptyString
    
    Capacity = Int(Capacity)
    
    If Capacity <= 0 Then _
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_NeedPosNum
    If Access < MemoryMappedFileAccess.ReadWrite Or Access > MemoryMappedFileAccess.ReadWriteExecute Then _
        Error.ArgumentOutOfRange "Access"
    
    If Access = MemoryMappedFileAccess.WriteOnly Then
        Set Handle = OpenCore(MapName, Access, True)
    Else
        Set Handle = OpenOrCreateCore(MapName, Capacity, Access)
    End If
    
    Set CreateOrOpen = Cor.NewMemoryMappedFile(Handle)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OpenOrCreateCore(ByRef MapName As String, ByVal Capacity As Currency, ByVal Access As MemoryMappedFileAccess) As SafeMemoryMappedFileHandle
    Dim Handle      As SafeMemoryMappedFileHandle
    Dim Capacity64  As Int64
    Dim LastError   As Long
    
    Capacity64 = CInt64(Capacity)
    Set Handle = SafeCreateFileMapping(-1, GetPageAccess(Access), Capacity64.HighPart, Capacity64.LowPart, MapName)
    LastError = Err.LastDllError
    
    If Handle.IsInvalid And (LastError <> ERROR_ALREADY_EXISTS) Then
        Error.Win32Error LastError
    End If
    
    Set OpenOrCreateCore = Handle
End Function

Private Function OpenCore(ByRef MapName As String, ByVal Access As MemoryMappedFileRights, ByVal IsCreateOrOpen As Boolean) As SafeMemoryMappedFileHandle
    Dim Handle      As SafeMemoryMappedFileHandle
    Dim LastError   As Long
    
    Set Handle = SafeOpenFileMapping(MapName, Access)
    LastError = Err.LastDllError
    
    If Handle.IsInvalid Then
        If IsCreateOrOpen And (LastError = ERROR_FILE_NOT_FOUND) Then _
            Error.Argument Argument_NewMMFWriteAccessNotAllowed, "Access"
        
        IOError.WinIOError LastError
    End If
    
    Set OpenCore = Handle
End Function

Private Function CreateCore(ByVal FileHandle As SafeFileHandle, ByRef MapName As String, ByVal Capacity As Currency, ByVal Access As MemoryMappedFileAccess) As SafeMemoryMappedFileHandle
    Dim Handle      As SafeMemoryMappedFileHandle
    Dim Capacity64  As Int64
    Dim LastError     As Long
    
    Capacity64 = CInt64(Capacity)
    Set Handle = SafeCreateFileMapping(FileHandle.DangerousGetHandle, GetPageAccess(Access), Capacity64.HighPart, Capacity64.LowPart, MapName)
    LastError = Err.LastDllError
    
    If Not Handle.IsInvalid And LastError = ERROR_ALREADY_EXISTS Then
        Handle.CloseHandle
        IOError.WinIOError LastError
    ElseIf Handle.IsInvalid Then
        IOError.WinIOError LastError
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

