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

Private Const vbPathSeparator               As Integer = vbSemiColonChar
Private Const vbVolumeSeparatorChar         As Integer = vbColonChar
Private Const vbDirectorySeparatorChar      As Integer = vbBackslashChar
Private Const vbAltDirectorySeparatorChar   As Integer = vbForwardSlashChar
Private Const vbVolumeSeparatorString       As String = ":"
Private Const vbDirectorySeparatorString    As String = "\"
Private Const vbAltDirectorySeparatorString As String = "/"

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
    AltDirectorySeparatorChar = vbAltDirectorySeparatorChar
End Property

''
' Returns the string for the alternative directory separator character.
'
' @return The string of the alternative directory separator "/".
' @remarks The alternate string is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparatorString() As String
    AltDirectorySeparatorString = "/"
End Property

''
' Returns the directory separator character for Windows.
'
' @return The directory separator character for Windows (92).
'
Public Property Get DirectorySeparatorChar() As Integer
    DirectorySeparatorChar = vbDirectorySeparatorChar
End Property

''
' Returns the directory separator string for Windows.
'
' @return The directory separator string for windows "\".
'
Public Property Get DirectorySeparatorString() As String
    DirectorySeparatorString = vbDirectorySeparatorString
End Property

''
' The character code used to separate strings in an environment variable.
'
' @return The environment variable string separator character code (59).
'
Public Property Get PathSeparator() As Integer
    PathSeparator = vbPathSeparator
End Property

''
' The string used to seperate strings in an environment variable.
'
' @return The environment variable string separator ";".
'
Public Property Get PathSeparatorString() As String
    PathSeparatorString = ";"
End Property

''
' The character code used to separate volume labels from path information.
'
' @return The character code for the volume separator (58).
'
Public Property Get VolumeSeparatorChar() As Integer
    VolumeSeparatorChar = vbVolumeSeparatorChar
End Property

''
' Returns the string version of the volume separator.
'
' @return The string version of the volume separator ":".
'
Public Property Get VolumeSeparatorString() As String
    VolumeSeparatorString = vbVolumeSeparatorString
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
    
    If Not CString.IsNull(Extension) Then
        If Len(Extension) = 0 Then
            Extension = "."
        ElseIf Asc(Extension) <> vbPeriodChar Then
            Extension = "." & Extension
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
    
    FixupPath Path
    
    With FirstTwoChars(Path)
        If .LoWord = vbDirectorySeparatorChar Then
            IsPathRooted = True
        ElseIf .HiWord = vbVolumeSeparatorChar Then
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
    VerifyPath Path
    
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
    VerifyPath Path1
    
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
    
    If IsSeparator(LastChar(Path1)) Then
        Combine = Path1 & Path2
    Else
        Combine = Path1 & Me.DirectorySeparatorString & Path2
    End If
End Function

Private Function IsSeparator(ByVal Char As Integer) As Boolean
    IsSeparator = Char = vbDirectorySeparatorChar Or _
                  Char = vbAltDirectorySeparatorChar Or _
                  Char = vbVolumeSeparatorChar
End Function

''
' Returns the directory portion of a file path.
'
' @param Path The path to parse the directory from.
' @return The directory parsed from the path.
' @remarks The function does not verify that the directory exists.
'
Public Function GetDirectoryName(ByVal Path As String) As String
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    FixupPath Path
    
    Dim MinSlashes As Long
    With FirstTwoChars(Path)
        If (.LoWord = vbDirectorySeparatorChar) And (.HiWord = vbDirectorySeparatorChar) Then
            MinSlashes = 4
        ElseIf .HiWord = vbVolumeSeparatorChar Then
            If (.LoWord = vbVolumeSeparatorChar) Or (InStr(3, Path, vbVolumeSeparatorString) > 0) Then _
                Error.Argument Argument_InvalidPathFormat, "Path"
            
            Select Case Asc(Path)
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    If Len(Path) > 3 Then
                        MinSlashes = 1
                    End If
                    
                Case Else
                    Error.Argument Argument_InvalidPathFormat, "Path"
            End Select
            
        Else
            MinSlashes = 1
        End If
    End With
    
    Dim PathLength As Long
    If (MinSlashes > 0) And (CString.CharCount(Path, vbDirectorySeparatorChar) >= MinSlashes) Then
        PathLength = InStrRev(Path, vbDirectorySeparatorString)
        If PathLength = 3 Then
            If Mid$(Path, 2, 1) <> vbVolumeSeparatorString Then
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
    Index = InStrRev(Path, ".")
    If (Index > 0) And (Index < Len(Path)) Then
        GetExtension = Mid$(Path, Index)
    End If
End Function

''
' Returns the filename portion of the specified path.
'
' @param Path The path to retrieve the filename from.
' @return The filetime portion of the specified path.
' @remarks This function does not verify that the file exists.
'
Public Function GetFileName(ByVal Path As String) As String
    VerifyPath Path
    
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
    Index = InStrRev(Path, ".")
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
    Const DoubleBackSlash As Long = &H5C005C
    
    Path = CString.Trim(Path)
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
        
    FixupPath Path
    
    Dim Ptr As Long
    Ptr = StrPtr(Path)
    
    ' quickly check the first two charcters for possible UNC path.
    ' convert first two characters into a long.
    If MemLong(Ptr) = DoubleBackSlash Then
        VerifyUNCPathFormat Path
    End If
    
    ' convert first two characters into two integers.
    With MemDWord(Ptr)
        If .LoWord = vbVolumeSeparatorChar Then _
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidPathFormat), "Path")

        ' check if absolute path ( a:, b:, ...) is being specified.
        If .HiWord = vbVolumeSeparatorChar Then
            Select Case .LoWord
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                Case Else
                    Throw Cor.NewArgumentException("Incorrect absolute path information.", "Path")
            End Select
        End If
    End With
    
    ' make sure no more than 1 volume separator exists.
    If InStr(3, Path, vbVolumeSeparatorString) > 0 Then _
        Throw Cor.NewNotSupportedException("Path format is not supported.")
    
    Dim PathParts() As String
    PathParts = Split(Path, vbDirectorySeparatorString)
    
    Dim i As Long
    For i = 0 To UBound(PathParts)
        Select Case PathParts(i)
            Case "..", "."
                ' do nothing
            Case Else
                PathParts(i) = CString.TrimEnd(PathParts(i), ".")
        End Select
    Next i
    
    Path = Join(PathParts, vbDirectorySeparatorString)
    
    Dim OutputBuffer As String
    OutputBuffer = String$(Me.MaxPathLength, 0)
    
    Dim Size As Long
    Size = Api.GetFullPathName(Path, Len(OutputBuffer), OutputBuffer, 0)
    
    If Size = 0 Then _
        Throw Cor.NewPathTooLongException("The qualified path is longer than the maximum of 260 characters.")
    
    GetFullPath = Left$(OutputBuffer, Size)
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
        Error.Argument Argument_EmptyPath, "Path"
    
    FixupPath Path
    
    With MemDWord(StrPtr(Path))
        If .LoWord = vbDirectorySeparatorChar Then
            Select Case .HiWord
                Case vbDirectorySeparatorChar
                    GetPathRoot = GetUNCPathRoot(Path)
                    Exit Function
                
                Case vbVolumeSeparatorChar
                    Error.Argument Argument_InvalidPathFormat, "Path"
                
                Case Else
                    GetPathRoot = vbDirectorySeparatorString
                    Exit Function
            End Select
        End If
        
        If .HiWord = vbVolumeSeparatorChar Then
            Select Case .LoWord
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    GetPathRoot = Left$(Path, 2)
                    
                    If Len(Path) > 2 Then
                        GetPathRoot = GetPathRoot + vbDirectorySeparatorString
                    End If
                    
                Case Else
                    Error.Argument Argument_InvalidPathFormat, "Path"
                    
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
    
    If Api.GetTempFileName(GetTempPath, "tmp", 0, FileName) = BOOL_FALSE Then IOError Err.LastDllError
    
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
    Size = Api.GetTempPath(Len(PathName), PathName)
    If Size = 0 Or Size > Me.MaxPathLength Then
        IOError Err.LastDllError
    End If

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
'   Friend
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub VerifyPath(ByRef Path As String, Optional ByVal FixSlashes As Boolean = False)
    If Len(Path) > MAX_PATH Then _
        Throw New PathTooLongException
    
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        Select Case mBuffer.Data(i)
            Case vbQuoteChar, vbLessThanChar, vbGreaterThanChar, vbPipeChar, Is < vbSpaceChar
                Error.Argument Argument_InvalidCharsInPath, "Path"
                
            Case vbAltDirectorySeparatorChar
                If FixSlashes Then
                    mBuffer.Data(i) = vbDirectorySeparatorChar
                End If
        End Select
    Next

    mBuffer.SA.pvData = vbNullPtr
End Sub

Friend Sub CheckValidPath(ByRef Path As String)
    If HasIllegalCharacters(Path) Then
        Error.Argument Argument_InvalidCharsInPath
    End If
End Sub

Friend Function HasIllegalCharacters(ByRef Path As String) As Boolean
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        Select Case mBuffer.Data(i)
            Case vbQuoteChar, vbLessThanChar, vbGreaterThanChar, vbPipeChar, Is < vbSpaceChar
                HasIllegalCharacters = True
                Exit For
        End Select
    Next
    
    mBuffer.SA.pvData = vbNullPtr
End Function

Friend Sub InplaceFixSlashes(ByRef Path As String)
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        If mBuffer.Data(i) = vbAltDirectorySeparatorChar Then
            mBuffer.Data(i) = vbDirectorySeparatorChar
        End If
    Next
    
    mBuffer.SA.pvData = vbNullPtr
End Sub

Friend Sub CheckValidPathAndFixSlashes(ByRef Path As String)
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        Select Case mBuffer.Data(i)
            Case vbQuoteChar, vbLessThanChar, vbGreaterThanChar, vbPipeChar, Is < vbSpaceChar
                Error.Argument Argument_InvalidCharsInPath, "Path"
                
            Case vbAltDirectorySeparatorChar
                mBuffer.Data(i) = vbDirectorySeparatorChar
        End Select
    Next
    
    mBuffer.SA.pvData = vbNullPtr
End Sub


Friend Sub FixupPath(ByRef Path As String)
    VerifyPath Path, True
    
    mPath.Length = 0
    mPath.AppendString Path
    
    Dim i As Long
    Do
        i = mPath.Length
        mPath.Replace "\\", "\", 1, i - 1
    Loop While i > mPath.Length
    Do
        i = mPath.Length
        mPath.Replace "...\", "..\"
    Loop While i > mPath.Length
    
    Dim Parts() As String
    Parts = Split(mPath.ToString, "\")
    
    For i = 0 To UBound(Parts)
        If Len(Parts(i)) > 2 Then
            Parts(i) = CString.TrimEnd(Parts(i), ".")
        End If
    Next i
    
    Path = Join(Parts, vbDirectorySeparatorString)
End Sub

Friend Function InternalGetFullPath(ByVal Path As String) As String
    VerifyPath Path
    
    If Len(Path) = 0 Then
        InternalGetFullPath = Directory.GetCurrentDirectory
    Else
        InternalGetFullPath = Replace$(Path, vbAltDirectorySeparatorString, vbDirectorySeparatorString)
        If Asc(Right$(InternalGetFullPath, 1)) <> vbDirectorySeparatorChar Then
            InternalGetFullPath = InternalGetFullPath & vbDirectorySeparatorString
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
    Do While mBuffer.Data(pos) = vbDirectorySeparatorChar
        pos = pos + 1
    Loop
    
    Dim SlashCount As Long
    Dim SlashIndex As Long
    SlashIndex = pos
    Do While (SlashCount < 2) And (mBuffer.Data(SlashIndex) <> 0)
        If mBuffer.Data(SlashIndex) = vbDirectorySeparatorChar Then
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

    Dim Ch As Integer

    mBuffer.SA.pvData = StrPtr(Path)
    pos = 2
    Ch = mBuffer.Data(pos)
    Do While Ch <> 0
        Select Case Ch
            Case vbDirectorySeparatorChar, vbAltDirectorySeparatorChar
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
        Ch = mBuffer.Data(pos)
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
            Case vbPeriodChar
                Exit Do
            Case vbAltDirectorySeparatorChar, vbDirectorySeparatorChar, vbVolumeSeparatorChar
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
            Case vbDirectorySeparatorChar, vbAltDirectorySeparatorChar, vbVolumeSeparatorChar
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
        Size = Api.GetLongPathName(ShortPath, LongPath, Size)
    Loop While Size > Len(LongPath)
    
    GetLongPathName = Left$(LongPath, Size)
End Function

Private Function FirstTwoChars(ByRef s As String) As DWord
    If LenB(s) > 0 Then
        FirstTwoChars = MemDWord(StrPtr(s))
    End If
End Function

Private Function LastChar(ByRef s As String) As Integer
    Dim Length As Long
    Length = Len(s)
    
    If Length > 0 Then
        Dim LastCharIndex As Long
        LastCharIndex = Length - 1
        LastChar = MemWord(StrPtr(s) + LastCharIndex * 2)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mInvalidPathChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F)
    mInvalidFileNameChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H3A, &H2A, &H3F, &H5C, &H2F)
    
    InitWordBuffer mBuffer, 0, &H7FFFFFFF
End Sub
