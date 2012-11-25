VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
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
' Module: Path
'

''
' Provides functions to locate and format path information.
'
' @remarks All members of Path are static. To access them, use the
' Path.* convention.
' <pre>
' Dim p As String
' p = Path.Combine(App.Path, "file.txt")
' </pre>
'
Option Explicit

Private Const NOT_FOUND             As Long = &H0
Private Const vbBackSlashBackSlash  As Long = &H5C005C



Private mInvalidPathChars()     As Integer
Private mInvalidFileNameChars() As Integer
Private mBuffer                 As WordBuffer
Private mPath                   As New StringBuilder


''
' Returns the maximum length for a path.
'
' @return Tha maximum length allowed for a path.
' @remarks Returns 32767 on Windows NT machines, 260 for Win9x machines.
'
Public Property Get MaxPathLength() As Long
    MaxPathLength = IIf(Environment.IsNT, MAX_PATH_W, MAX_PATH_A)
End Property

''
' Returns the character code for the alternative directory separator character.
'
' @return The character code of the alternative directory separator (47).
' @remarks The alternate character is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparatorChar() As Integer
    AltDirectorySeparatorChar = vbForwardSlash
End Property

''
' Returns the string for the alternative directory separator character.
'
' @return The string of the alternative directory separator "/".
' @remarks The alternate string is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparatorString() As String
    AltDirectorySeparatorString = vbForwardSlashS
End Property

''
' Returns the directory separator character for Windows.
'
' @return The directory separator character for Windows (92).
'
Public Property Get DirectorySeparatorChar() As Integer
    DirectorySeparatorChar = vbBackSlash
End Property

''
' Returns the directory separator string for Windows.
'
' @return The directory separator string for windows "\".
'
Public Property Get DirectorySeparatorString() As String
    DirectorySeparatorString = vbBackSlashS
End Property

''
' Returns an array of characters that are invalid in a Windows path string.
'
' @return An Integer array representing the character code of the the invalid characters.
'
Public Property Get InvalidPathChars() As Integer()
    InvalidPathChars = mInvalidPathChars
End Property

''
' The character code used to separate strings in an environment variable.
'
' @return The environment variable string separator character code (59).
'
Public Property Get PathSeparator() As Integer
    PathSeparator = vbSemiColon
End Property

''
' The string used to seperate strings in an environment variable.
'
' @return The environment variable string separator ";".
'
Public Property Get PathSeparatorString() As String
    PathSeparatorString = vbSemiColonS
End Property

''
' The character code used to separate volume labels from path information.
'
' @return The character code for the volume separator (58).
'
Public Property Get VolumeSeparatorChar() As Integer
    VolumeSeparatorChar = vbColon
End Property

''
' Returns the string version of the volume separator.
'
' @return The string version of the volume separator ":".
'
Public Property Get VolumeSeparatorString() As String
    VolumeSeparatorString = vbColonS
End Property

''
' Replaces the extension of a filename with a new extension.
'
' @param Path The filename with an optional path.
' @param Extension The new extension to change the filename's to.
' @return The filename with the new extension, including a path if supplied.
' @remarks Passing in an <i>Extension</i> of <i>vbNullString</i> will remove
' any extension and period from the path.
'
Public Function ChangeExtension(ByVal Path As String, ByVal Extension As String) As String
    If Len(Path) = 0 Then Exit Function
    
    VerifyPath Path
    
    If Not cString.IsNull(Extension) Then
        ' replace or append extension.
        If Len(Extension) = 0 Then
            Extension = vbPeriodS
        ElseIf Asc(Extension) <> vbPeriod Then
            Extension = vbPeriodS & Extension
        End If
    End If
    
    Dim PeriodIndex As Long
    PeriodIndex = GetExtensionCharIndex(Path)
    If PeriodIndex = NOT_FOUND Then
        ChangeExtension = Path & Extension
    Else
        ChangeExtension = Left$(Path, PeriodIndex - 1) & Extension
    End If
End Function

''
' Returns an indication if the path is rooted or not.
'
' @param Path The path to check if is rooted.
' @return The indication of the path being rooted.
' @remarks <p>A rooted path is one that starts with a volume name,
' UNC server name, or directory separator. Such paths as C:\MyDir\File,
' \\Server\Share\File, and \MyFile are rooted.</p>
' <p>This function does not verify that the path actually exists.</p>
'
Public Function IsPathRooted(ByVal Path As String) As Boolean
    If Len(Path) = 0 Then Exit Function
    
    Call FixupPath(Path)
    
    With MemDWord(StrPtr(Path))
        If .LoWord = vbBackSlash Then
            IsPathRooted = True
        ElseIf .HiWord = vbColon Then
            IsPathRooted = True
        End If
    End With
End Function

''
' Returns if a filename has an extension.
'
' @param Path A filename and optional path to check for an extension.
' @return An indication of the filename having an extension.
' @remarks <p>The path string is search for a '.' that separates the
' the extension from the filename itself. If any of the directory separator
' characters or volume separator is encountered before a '.', then the
' filename is considered not to have an extension.
'
Public Function HasExtension(ByVal Path As String) As Boolean
    Call VerifyPath(Path)
    
    Select Case GetExtensionCharIndex(Path)
        Case Len(Path): Exit Function
        Case Is > 0: HasExtension = True
    End Select
End Function

''
' Combintes two paths into a single path.
'
' @param Path1 The left-hand side of the two paths to be joined.
' @param Path2 The right-hand side of the two paths to be joined.
' @return The joined paths.
' @remarks <p>If Path1 is an empty string, then Path2 is returned as is. If Path2
' is empty, but not Path1, then Path1 is returned as is.</p>
' <p>If Path2 is rooted (starts with a volume, UNC identifier or directory
' separator), then Path2 is returned, regardless of Path1.</p>
' <p>If Path1 does not end with a directory separator, one will be inserted.</p>
' <p>This function does not verify that the paths exist.</p>
'
Public Function Combine(ByVal Path1 As String, ByVal Path2 As String) As String
    Call VerifyPath(Path1)
    
    ' This will verify Path2 as well.
    If IsPathRooted(Path2) Then
        Combine = Path2
        Exit Function
    End If
    
    If Len(Path1) = 0 Then
        Combine = Path2
        Exit Function
    ElseIf Len(Path2) = 0 Then
        Combine = Path1
        Exit Function
    End If
    
    mBuffer.SA.pvData = StrPtr(Path1)
    Select Case mBuffer.Data(Len(Path1) - 1)
        Case vbBackSlash, vbForwardSlash, vbColon
            Combine = Path1 & Path2
        Case Else
            Combine = Path1 & vbBackSlashS & Path2
    End Select
End Function

''
' Returns the directory portion of a file path.
'
' @param Path The path to parse the directory from.
' @return The directory parsed from the path.
' @remarks The function does not verify that the directory exists.
'
Public Function GetDirectoryName(ByVal Path As String) As String
    Path = cString.Trim(Path)
    
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
        
    Call FixupPath(Path)
    
    Dim MinSlashes As Long
    With MemDWord(StrPtr(Path))
        If (.LoWord = vbBackSlash) And (.HiWord = vbBackSlash) Then
            MinSlashes = 4
            
        ElseIf .HiWord = vbColon Then
            If (.LoWord = vbColon) Or (InStr(3, Path, ":") > 0) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
            
            Select Case Asc(Path)
                Case vbLowerA To vbLowerZ, vbUpperA To vbUpperZ
                    If Len(Path) > 3 Then MinSlashes = 1
                
                Case Else: Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
            End Select
            
        Else
            MinSlashes = 1
        End If
    End With
    
    Dim PathLength As Long
    If (MinSlashes > 0) And (cString.CharCount(Path, vbBackSlash) >= MinSlashes) Then
        PathLength = InStrRev(Path, vbBackSlashS)
        If PathLength = 3 Then
            If Mid$(Path, 2, 1) <> ":" Then
                PathLength = PathLength - 1
            End If
        Else
            PathLength = PathLength - 1
        End If
    End If
    
    GetDirectoryName = Left$(Path, PathLength)
End Function

''
' Returns the extension of a filename.
'
' @param Path The filename to retrieve the extension of.
' @return The extension of the filename
'
Public Function GetExtension(ByVal Path As String) As String
    Path = GetFileName(Path)
    
    Dim Index As Long
    Index = InStrRev(Path, vbPeriodS)
    If (Index > 0) And (Index < Len(Path)) Then GetExtension = Mid$(Path, Index)
End Function

''
' Returns the filename portion of the specified path.
'
' @param Path The path to retrieve the filename from.
' @return The filetime portion of the specified path.
' @remarks This function does not verify that the file exists.
'
Public Function GetFileName(ByVal Path As String) As String
    Call VerifyPath(Path)
    
    GetFileName = Mid$(Path, GetDirectoryCharIndex(Path) + 1)
End Function

''
' Returns the filename portion of a path without the filename extension.
'
' @param Path The path to parse the filename from.
' @return A filename without the extension present.
' @remarks This function does not verify the file exists.
'
Public Function GetFileNameWithoutExtension(ByVal Path As String) As String
    Path = GetFileName(Path)
    
    Dim Index As Long
    Index = InStrRev(Path, vbPeriodS)
    If Index > 0 Then Path = Left$(Path, Index - 1)
    
    GetFileNameWithoutExtension = Path
End Function

''
' Creates a full path to the specified filename.
'
' @param Path The filename to prepend any path to.
' @return A fully qualified path to a filename.
' @remarks The filename is prepended with the current directory to
' give a full path. The file does not need to exists, nor is the
' existence verified.
' <p>If <i>Path</i> is already rooted, then <i>Path</i> is returned.</p>
'
Public Function GetFullPath(ByVal Path As String) As String
    Path = cString.Trim(Path)
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
    
    Call FixupPath(Path)
    
    Dim ptr As Long
    ptr = StrPtr(Path)
    
    ' quickly check the first two charcters for possible UNC path.
    ' convert first two characters into a long.
    If MemLong(ptr) = vbBackSlashBackSlash Then Call VerifyUNCPathFormat(Path)
    
    ' convert first two characters into two integers.
    With MemDWord(ptr)
        ' check first character being ':'
        If .LoWord = vbColon Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")

        ' check if absolute path ( a:, b:, ...) is being specified.
        If .HiWord = vbColon Then
            Select Case .LoWord
                Case vbLowerA To vbLowerZ, vbUpperA To vbUpperZ
                Case Else
                    Throw Cor.NewArgumentException("Incorrect absolute path information.", "Path")
            End Select
        End If
    End With
    
    ' make sure no more than 1 colon exists.
    If InStr(3, Path, ":") > 0 Then _
        Throw Cor.NewNotSupportedException("Path format is not supported.")
    
    Dim PathParts() As String
    PathParts = Split(Path, vbBackSlashS)
    
    Dim i As Long
    For i = 0 To UBound(PathParts)
        Select Case PathParts(i)
            Case "..", "."
                ' do nothing
            Case Else
                PathParts(i) = cString.TrimEnd(PathParts(i), vbPeriodS)
        End Select
    Next i
    
    Path = Join(PathParts, vbBackSlashS)
    
    Dim Ret As String
    Ret = String$(Me.MaxPathLength, 0)
    
    Dim Size As Long
    Size = API.GetFullPathName(Path, Len(Ret), Ret, 0)
    
    If Size = 0 Then _
        Throw Cor.NewPathTooLongException("The qualified path is longer than the maximum of 260 characters.")
    
    GetFullPath = Left$(Ret, Size)
End Function

''
' Returns the root directory of the specified path.
'
' @param Path The path to retrieve the root directory from.
' @return The root directory of the specified path.
' @remarks <i>Path</i> is not checked if it exists. The root is
' derived from what information is in <i>Path</i>.
'
Public Function GetPathRoot(ByVal Path As String) As String
    If Len(Path) = 0 Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
    
    Call FixupPath(Path)
    
    With MemDWord(StrPtr(Path))
        If .LoWord = vbBackSlash Then
            Select Case .HiWord
                Case vbBackSlash
                    GetPathRoot = GetUNCPathRoot(Path)
                    Exit Function
                
                Case vbColon
                    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
                
                Case Else
                    GetPathRoot = vbBackSlashS
                    Exit Function
            End Select
        End If
        
        If .HiWord = vbColon Then
            Select Case .LoWord
                Case vbLowerA To vbLowerZ, vbUpperA To vbUpperZ
                    GetPathRoot = Left$(Path, 2)
                    
                    If Len(Path) > 2 Then
                        GetPathRoot = GetPathRoot + vbBackSlashS
                    End If
                    Exit Function
                    
                Case Else
                    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")
            End Select
        End If
    End With
End Function

''
' Returns the filename for a temporary file.
'
' @return The temporary file's filename.
' @remark This function creates a 0 byte temporary file of the returned name.
' The file is placed in the temporary path of the current user.
'
Public Function GetTempFileName() As String
    Dim FileName As String
    FileName = String$(Me.MaxPathLength, 0)
    
    If API.GetTempFileName(GetTempPath, "tmp", 0, FileName) = BOOL_FALSE Then IOError Err.LastDllError
    
    GetTempFileName = SysAllocString(StrPtr(FileName))
End Function

''
' Returns the path to the temporary folder for the current user.
'
' @return The temporary path for the current user.
'
Public Function GetTempPath() As String
    Dim PathName As String
    Dim Size As Long
    
    PathName = String$(Me.MaxPathLength, 0)
    Size = API.GetTempPath(Len(PathName), PathName)
    If Size = 0 Or Size > Me.MaxPathLength Then Call IOError(Err.LastDllError)

    GetTempPath = GetLongPathName(Left$(PathName, Size))
End Function

''
' Returns an array of invalid path characters.
'
Public Function GetInvalidPathChars() As Integer()
    GetInvalidPathChars = mInvalidPathChars
End Function

''
' Returns an array of invalid filename characters.
'
Public Function GetInvalidFileNameChars() As Integer()
    GetInvalidFileNameChars = mInvalidFileNameChars
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub VerifyPath(ByRef Path As String, Optional ByVal FixSlashes As Boolean = False)
    If Len(Path) > MAX_PATH Then _
        Throw New PathTooLongException
    
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        ' We check for invalid path characters with hardcoded values instead
        ' of looping through both the path to check and the list of invalid
        ' characters which would be much slower.
        Select Case mBuffer.Data(i)
            Case &H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F
                Throw Cor.NewArgumentException("Invalid characters in path.", "Path")
            Case vbForwardSlash
                If FixSlashes Then mBuffer.Data(i) = vbBackSlash
        End Select
    Next i
End Sub

Friend Sub FixupPath(ByRef Path As String)
    Call VerifyPath(Path, True)
    
    mPath.Length = 0
    Call mPath.AppendQuick(Path)
    
    Dim i As Long
    Do
        i = mPath.Length
        Call mPath.Replace("\\", "\", 1, i - 1)
    Loop While i > mPath.Length
    Do
        i = mPath.Length
        Call mPath.Replace("...\", "..\")
    Loop While i > mPath.Length
    
    Dim Parts() As String
    Parts = Split(mPath.ToString, "\")
    
    For i = 0 To UBound(Parts)
        If Len(Parts(i)) > 2 Then
            Parts(i) = cString.TrimEnd(Parts(i), ".")
        End If
    Next i
    
    Path = Join(Parts, vbBackSlashS)
End Sub

Friend Function InternalGetFullPath(ByVal Path As String) As String
    Call VerifyPath(Path)
    
    If Len(Path) = 0 Then
        InternalGetFullPath = Directory.GetCurrentDirectory
    Else
        InternalGetFullPath = Replace$(Path, vbForwardSlashS, vbBackSlashS)
        If Asc(Right$(InternalGetFullPath, 1)) <> vbBackSlash Then
            InternalGetFullPath = InternalGetFullPath & vbBackSlashS
        End If
    End If
End Function

Friend Function InternalGetFullRootedPath(ByVal Path As String) As String
    If Not IsPathRooted(Path) Then
        Path = GetFullPath(Path)
    End If
    InternalGetFullRootedPath = InternalGetFullPath(Path)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetUNCPathRoot(ByRef Path As String) As String
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim pos As Long
    pos = 2 ' skip the first 2 slashes.
    
    ' Skip backslashes after the first two
    ' until we get a character or the end.
    ' We don't need to check for a valid index
    ' because we'll hit the null terminator.
    Do While mBuffer.Data(pos) = vbBackSlash: pos = pos + 1: Loop
    
    Dim SlashCount As Long
    Dim SlashIndex As Long
    SlashIndex = pos
    Do While (SlashCount < 2) And (mBuffer.Data(SlashIndex) <> 0)
        If mBuffer.Data(SlashIndex) = vbBackSlash Then
            SlashCount = SlashCount + 1
        End If
        
        SlashIndex = SlashIndex + 1
    Loop
    
    If SlashCount = 2 Then SlashIndex = SlashIndex - 1 ' we don't want to include the 2nd slash.
    
    GetUNCPathRoot = "\\" & Mid$(Path, pos + 1, SlashIndex - pos)
End Function

Private Sub VerifyUNCPathFormat(ByRef Path As String)
    Dim pos             As Long
    Dim foundServer     As Boolean
    Dim foundShare      As Boolean
    Dim foundSeparator  As Boolean

    Dim ch As Integer

    mBuffer.SA.pvData = StrPtr(Path)
    pos = 2
    ch = mBuffer.Data(pos)
    Do While ch <> 0
        Select Case ch
            Case vbBackSlash, vbForwardSlash
                foundSeparator = True
            Case Else
                If Not foundServer Then
                    foundServer = True
                ElseIf Not foundShare And foundSeparator Then
                    foundShare = True
                    Exit Do
                End If
        End Select
        pos = pos + 1
        ch = mBuffer.Data(pos)
    Loop

    If (Not foundServer) Or (Not foundShare) Then _
        Throw Cor.NewArgumentException("UNC path must be in the format of \\server\share.")
End Sub

''
' Returns a 1-based index for the extension separator character
' in the string. Works like the InStrRev.
'
' @param Path The path to search for the extension separator character.
' @return A 1-based index of the separator character, or 0 if not found.
'
Private Function GetExtensionCharIndex(ByRef Path As String) As Long
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim Index As Long
    Index = Len(Path) - 1
    
    Do While Index >= 0
        Select Case mBuffer.Data(Index)
            Case vbPeriod: Exit Do
            Case vbForwardSlash, vbBackSlash, vbColon
                GetExtensionCharIndex = NOT_FOUND
                Exit Function
        End Select
        Index = Index - 1
    Loop
    
    GetExtensionCharIndex = Index + 1
End Function

''
' Returns a 1-based index for the last separator character
' in the string. Works like the InStrRev.
'
' @param Path The path to search for the last separator character.
' @return A 1-based index of the separator character, or 0 if not found.
'
Private Function GetDirectoryCharIndex(ByRef Path As String) As Long
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim Index As Long
    Index = Len(Path) - 1
    Do While Index >= 0
        Select Case mBuffer.Data(Index)
            Case vbBackSlash, vbForwardSlash, vbColon
                Exit Do
        End Select
        Index = Index - 1
    Loop
    
    GetDirectoryCharIndex = Index + 1
End Function

Private Function GetLongPathName(ByRef ShortPath As String) As String
    Dim LongPath    As String
    Dim Size        As Long
    
    Size = 1024
    Do
        LongPath = String$(Size, 0)
        Size = API.GetLongPathName(ShortPath, LongPath, Size)
    Loop While Size > Len(LongPath)
    
    GetLongPathName = Left$(LongPath, Size)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mInvalidPathChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F)
    mInvalidFileNameChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H3A, &H2A, &H3F, &H5C, &H2F)
    
    Call InitWordBuffer(mBuffer, 0, &H7FFFFFFF)
End Sub
