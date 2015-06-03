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

Private mInvalidPathChars()     As Integer
Private mInvalidFileNameChars() As Integer
Private mTrimEndChars()         As Integer
Private mBuffer                 As WordBuffer
Private mPath                   As New StringBuilder


Friend Property Get MaxPath() As Long
    MaxPath = MAX_PATH
End Property

Friend Property Get MaxLongPath() As Long
    MaxLongPath = MAX_LONG_PATH
End Property

''
' Returns the character code for the alternative directory separator character.
'
' @return The character code of the alternative directory separator (47).
' @remarks The alternate character is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparatorChar() As Integer
    AltDirectorySeparatorChar = corAltDirectorySeparatorChar
End Property

''
' Returns the string for the alternative directory separator character.
'
' @return The string of the alternative directory separator "/".
' @remarks The alternate string is usually found on systems such as Unix.
'
Public Property Get AltDirectorySeparator() As String
    AltDirectorySeparator = corAltDirectorySeparator
End Property

''
' Returns the directory separator character for Windows.
'
' @return The directory separator character for Windows (92).
'
Public Property Get DirectorySeparatorChar() As Integer
    DirectorySeparatorChar = corDirectorySeparatorChar
End Property

''
' Returns the directory separator string for Windows.
'
' @return The directory separator string for windows "\".
'
Public Property Get DirectorySeparator() As String
    DirectorySeparator = corDirectorySeparator
End Property

''
' The character code used to separate strings in an environment variable.
'
' @return The environment variable string separator character code (59).
'
Public Property Get PathSeparatorChar() As Integer
    PathSeparatorChar = corPathSeparatorChar
End Property

''
' The string used to seperate strings in an environment variable.
'
' @return The environment variable string separator ";".
'
Public Property Get PathSeparator() As String
    PathSeparator = corPathSeparator
End Property

''
' The character code used to separate volume labels from path information.
'
' @return The character code for the volume separator (58).
'
Public Property Get VolumeSeparatorChar() As Integer
    VolumeSeparatorChar = corVolumeSeparatorChar
End Property

''
' Returns the string version of the volume separator.
'
' @return The string version of the volume separator ":".
'
Public Property Get VolumeSeparator() As String
    VolumeSeparator = corVolumeSeparator
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
Public Function ChangeExtension(ByRef Path As String, ByRef Extension As String) As String
    If Len(Path) = 0 Then
        Exit Function
    End If
    
    CheckValidPathChars Path
    
    Dim PeriodIndex As Long
    PeriodIndex = GetExtensionCharIndex(Path)
    
    If PeriodIndex > 0 Then
        ChangeExtension = Left$(Path, PeriodIndex - 1)
    Else
        ChangeExtension = Path
    End If
    
    If Not CorString.IsNull(Extension) Then
        If Len(Extension) = 0 Then
            ChangeExtension = ChangeExtension & "."
        ElseIf AscW(Extension) = corPeriodChar Then
            ChangeExtension = ChangeExtension & Extension
        Else
            ChangeExtension = ChangeExtension & "." & Extension
        End If
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
Public Function IsPathRooted(ByRef Path As String) As Boolean
    If Len(Path) = 0 Then Exit Function
    
    CheckValidPathChars Path
    
    With FirstTwoChars(Path)
        If IsDirectorySeparator(.LoWord) Then
            IsPathRooted = True
        ElseIf .HiWord = corVolumeSeparatorChar Then
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
Public Function HasExtension(ByRef Path As String) As Boolean
    CheckValidPathChars Path
    
    Dim Index As Long
    Index = GetExtensionCharIndex(Path)
    
    HasExtension = Index > 0 And Index < Len(Path)
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
Public Function Combine(ByRef Path1 As String, ByRef Path2 As String) As String
    CheckValidPathChars Path1
    
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
    
    If IsSeparator(CorString.LastChar(Path1)) Then
        Combine = Path1 & Path2
    Else
        Combine = Path1 & corDirectorySeparator & Path2
    End If
End Function

''
' Returns the directory portion of a file path.
'
' @param Path The path to parse the directory from.
' @return The directory parsed from the path.
' @remarks The function does not verify that the directory exists.
'
Public Function GetDirectoryName(ByRef Path As String) As String
    If Len(Path) = 0 Then _
        Throw Error.Argument(Argument_EmptyPath, "Path")
    
    CheckValidPathChars Path
    
    Dim NormalizedPath As String
    NormalizedPath = NormalizePath(Path, False)
    
    Dim PathLength As Long
    Dim RootLength As Long
    RootLength = GetRootLength(NormalizedPath)
    
    If RootLength < Len(NormalizedPath) Then
        PathLength = InStrRev(NormalizedPath, corDirectorySeparator)
        
        If PathLength = 3 Then
            If RootLength <> 3 Then
                PathLength = PathLength - 1
            End If
        ElseIf PathLength > 0 Then
            PathLength = PathLength - 1
        End If
    End If
    
    GetDirectoryName = Left$(NormalizedPath, PathLength)
End Function

''
' Returns the extension of a filename.
'
' @param Path The filename to retrieve the extension of.
' @return The extension of the filename
'
Public Function GetExtension(ByRef Path As String) As String
    Dim Index As Long
    Index = GetExtensionCharIndex(Path)
    If Index > 0 And Index < Len(Path) Then
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
Public Function GetFileName(ByRef Path As String) As String
    CheckValidPathChars Path
    
    Dim Index As Long
    Index = GetDirectoryCharIndex(Path)
    
    If Index > 0 Then
        GetFileName = Mid$(Path, Index + 1)
    Else
        GetFileName = Path
    End If
End Function

''
' Returns the filename portion of a path without the filename extension.
'
' @param Path The path to parse the filename from.
' @return A filename without the extension present.
' @remarks This function does not verify the file exists.
'
Public Function GetFileNameWithoutExtension(ByRef Path As String) As String
    Dim FileName As String
    FileName = GetFileName(Path)
    
    Dim Index As Long
    Index = InStrRev(FileName, ".")
    If Index > 0 Then
        GetFileNameWithoutExtension = Left$(FileName, Index - 1)
    Else
        GetFileNameWithoutExtension = FileName
    End If
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
Public Function GetFullPath(ByRef Path As String) As String
    Dim FullPath As String
    FullPath = NormalizePath(Path, True)
    
    If InStr(3, FullPath, corVolumeSeparator) >= 3 Then _
        Throw Error.NotSupported(Argument_PathFormatNotSupported)
        
    GetFullPath = FullPath
End Function

''
' Returns the root directory of the specified path.
'
' @param Path The path to retrieve the root directory from.
' @return The root directory of the specified path.
' @remarks <i>Path</i> is not checked if it exists. The root is
' derived from what information is in <i>Path</i>.
'
Public Function GetPathRoot(ByRef Path As String) As String
    Dim NormalizedPath As String
    NormalizedPath = NormalizePath(Path, False)
    GetPathRoot = Left$(NormalizedPath, GetRootLength(NormalizedPath))
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
    FileName = String$(MaxPath, 0)
    
    If GetTempFileNameW(GetTempPath, "tmp", 0, FileName) = BOOL_FALSE Then
        IOError.WinIOError Err.LastDllError
    End If
    
    GetTempFileName = SysAllocString(StrPtr(FileName))
End Function

''
' Returns the path to the temporary folder for the current user.
'
' @return The temporary path for the current user.
'
Public Function GetTempPath() As String
    Dim PathName    As String
    Dim Size        As Long
    
    PathName = String$(MaxPath, 0)
    Size = GetTempPathW(MaxPath, PathName)
    
    If Size = 0 Or Size > MaxPath Then
        IOError.WinIOError Err.LastDllError
    End If

    GetTempPath = GetFullPath(Left$(PathName, Size))
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


Friend Sub VerifyPath(ByRef Path As String, Optional ByVal FixSlashes As Boolean = False)
    If Len(Path) > MAX_PATH Then _
        Throw New PathTooLongException
    
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        Select Case mBuffer.Data(i)
            Case corQuoteChar, corLessThanChar, corGreaterThanChar, corPipeChar, Is < corSpaceChar
                Throw Error.Argument(Argument_InvalidPathChars, "Path")
                
            Case corAltDirectorySeparatorChar
                If FixSlashes Then
                    mBuffer.Data(i) = corDirectorySeparatorChar
                End If
        End Select
    Next

    mBuffer.SA.pvData = vbNullPtr
End Sub

Friend Sub CheckValidPathChars(ByRef Path As String)
    If HasIllegalCharacters(Path) Then
        Throw Error.Argument(Argument_InvalidPathChars)
    End If
End Sub

Friend Function HasIllegalCharacters(ByRef Path As String) As Boolean
    Dim PathChars() As Integer
    PathChars = AllocChars(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        Select Case PathChars(i)
            Case corQuoteChar, corLessThanChar, corGreaterThanChar, corPipeChar, Is < corSpaceChar
                HasIllegalCharacters = True
                Exit For
        End Select
    Next
    
    FreeChars PathChars
End Function

Friend Sub CheckValidPathAndFixSlashes(ByRef Path As String)
    mBuffer.SA.pvData = StrPtr(Path)
    
    Dim i As Long
    For i = 0 To Len(Path) - 1
        Select Case mBuffer.Data(i)
            Case corQuoteChar, corLessThanChar, corGreaterThanChar, corPipeChar, Is < corSpaceChar
                Throw Error.Argument(Argument_InvalidPathChars, "Path")
                
            Case corAltDirectorySeparatorChar
                mBuffer.Data(i) = corDirectorySeparatorChar
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
            Parts(i) = CorString.TrimEnd(Parts(i), ".")
        End If
    Next i
    
    Path = Join(Parts, corDirectorySeparator)
End Sub

Friend Function NormalizePath(ByRef Path As String, ByVal FullCheck As Boolean) As String
    Dim NewBuffer   As New PathHelper
    Dim PathChars() As Integer
    Dim Index       As Long
    Dim MaxIndex    As Long
    
    If FullCheck Then
        Dim CleanPath As String
        CleanPath = CorString.TrimEnd(Path, mTrimEndChars)
        If HasIllegalCharacters(CleanPath) Then _
            Throw Error.Argument(Argument_InvalidPathChars)
        
        PathChars = AllocChars(CleanPath)
    Else
        PathChars = AllocChars(Path)
    End If
    
    If Not PrepareForUNCPath(PathChars, Index, NewBuffer) Then
        PrepareForVolumeLabel PathChars, Index, NewBuffer
    End If
    
    MaxIndex = UBound(PathChars)
    Do While Index <= MaxIndex
        If ProcessPreSigCharSegment(PathChars, NewBuffer, FullCheck, Index) Then
            ProcessSigCharSegment PathChars, NewBuffer, FullCheck, Index
        End If
    Loop
    
    FreeChars PathChars
    
    If NewBuffer.Length = 0 Then _
        Throw Error.Argument(Arg_PathIllegal)
    
    Dim MaxUNCPathCheck As Long
    If FullCheck Then
        If NewBuffer.StartsWith("http:") Or NewBuffer.StartsWith("file:") Then _
            Throw Error.Argument(Argument_PathUriFormatNotSupported)
    
        MaxUNCPathCheck = NewBuffer.GetFullPathName
    Else
        MaxUNCPathCheck = 1
    End If
    
    If MaxUNCPathCheck <> 0 Then
        If NewBuffer.IsUNCPath Then
            VerifyUNCPathFormat NewBuffer, MaxUNCPathCheck
        End If
    End If
    
    NormalizePath = NewBuffer.ToString
End Function

Friend Function IsDirectorySeparator(ByVal Ch As Integer) As Boolean
    IsDirectorySeparator = Ch = corDirectorySeparatorChar Or _
                           Ch = corAltDirectorySeparatorChar
End Function

Private Function PrepareForUNCPath(ByRef PathChars() As Integer, ByRef Index As Long, ByVal NewBuffer As PathHelper) As Boolean
    If UBound(PathChars) > 0 Then
        If PathChars(0) = corDirectorySeparatorChar Or PathChars(0) = corAltDirectorySeparatorChar Then
            NewBuffer.Append corDirectorySeparatorChar
            Index = 1
            PrepareForUNCPath = True
        End If
    End If
End Function

Private Sub PrepareForVolumeLabel(ByRef PathChars() As Integer, ByRef Index As Long, ByVal NewBuffer As PathHelper)
    If UBound(PathChars) > 0 Then
        If PathChars(1) = corVolumeSeparatorChar Then
            Select Case PathChars(0)
                Case corLowerAChar To corLowerZChar, corUpperAChar To corUpperZChar
                    NewBuffer.SetVolumeLabel PathChars(0)
                    Index = 2
                Case Else
                    Throw Error.Argument(Arg_PathIllegal)
            End Select
        End If
    End If
End Sub


Private Function ProcessPreSigCharSegment(ByRef PathChars() As Integer, ByVal NewBuffer As PathHelper, ByVal FullCheck As Boolean, ByRef Index As Long) As Boolean
    Dim pos             As Long
    Dim Ch              As Integer
    Dim MaxIndex        As Long
    Dim HasSpaces       As Boolean
    Dim HasPeriods      As Boolean
    Dim FoundSigChar    As Boolean
    Dim FoundSlash      As Boolean
    Dim LastPeriodIndex As Long
    Dim LastSpaceIndex  As Long
    
    MaxIndex = UBound(PathChars)
    
    pos = Index
    Do While pos <= MaxIndex
        Ch = PathChars(pos)
        
        Select Case Ch
            Case corPeriodChar
                HasPeriods = True
                LastPeriodIndex = pos
            Case corSpaceChar
                HasSpaces = True
                LastSpaceIndex = pos
            Case corDirectorySeparatorChar, corAltDirectorySeparatorChar
                If pos > Index Then
                    Exit Do
                End If
                
                If Not FoundSlash Then
                    NewBuffer.Append corDirectorySeparatorChar
                    FoundSlash = True
                End If
                Index = Index + 1
            Case Else
                FoundSigChar = True
                Exit Do
        End Select
        
        pos = pos + 1
    Loop
    
    Dim LastCharIndex As Long
    If FoundSigChar Then
        LastCharIndex = pos - 1
    Else
        If HasSpaces And ((LastSpaceIndex < LastPeriodIndex) Or (Not HasPeriods)) Then
            FreeChars PathChars
            Throw Error.Argument(Arg_PathIllegal)
        End If
    
        LastCharIndex = MathExt.Min(Index + 1, LastPeriodIndex)
    End If
    
    Dim j As Long
    For j = Index To LastCharIndex
        NewBuffer.Append PathChars(j)
    Next
    
    Index = pos
    ProcessPreSigCharSegment = FoundSigChar
End Function

Private Sub ProcessSigCharSegment(ByRef PathChars() As Integer, ByVal NewBuffer As PathHelper, ByVal FullCheck As Boolean, ByRef Index As Long)
    Dim pos         As Long
    Dim Ch          As Integer
    Dim MaxIndex    As Long
    Dim LastSigChar As Long
    Dim TryExpand   As Boolean
    
    MaxIndex = UBound(PathChars)
    
    pos = Index
    Do While pos <= MaxIndex
        Ch = PathChars(pos)
        
        Select Case Ch
            Case corTildaChar
                TryExpand = True
            Case corDirectorySeparatorChar, corAltDirectorySeparatorChar
                Exit Do
            Case corPeriodChar, corSpaceChar
            Case Else
                LastSigChar = pos
        End Select
    
        pos = pos + 1
    Loop
    
    Dim j As Long
    For j = Index To LastSigChar
        NewBuffer.Append PathChars(j)
    Next
    
    If TryExpand Then
        NewBuffer.TryExpandShortFileName
    End If
    
    Index = pos
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mInvalidPathChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F)
    mInvalidFileNameChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H3A, &H2A, &H3F, &H5C, &H2F)
    mTrimEndChars = Cor.NewIntegers(&H9, &HA, &HB, &HC, &HD, &H20, &H85, &HA0)
    
    InitWordBuffer mBuffer, 0, &H7FFFFFFF
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRootLength(ByRef Path As String) As Long
    Dim PathChars() As Integer
    PathChars = AllocChars(Path)
    
    Dim Result As Long
    Dim Length As Long
    Length = Len(Path)
    
    If Length > 0 Then
        With AsDWord(PathChars(0))
            If IsDirectorySeparator(.LoWord) Then
                If IsDirectorySeparator(.HiWord) Then
                    ' find unc server->share length
                    Dim SlashCount As Long
                    Dim i As Long
                    i = 2
                    
                    Do While i < Length
                        If IsDirectorySeparator(PathChars(i)) Then
                            SlashCount = SlashCount + 1
                            
                            If SlashCount = 2 Then
                                Exit Do
                            End If
                        End If
                        
                        i = i + 1
                    Loop
                    
                    Result = i
                Else
                    Result = 1
                End If
            ElseIf .HiWord = corVolumeSeparatorChar Then
                Result = 2
                
                ' check for following directory separator
                If Length > 2 Then
                    If PathChars(2) = corDirectorySeparatorChar Then
                        Result = 3
                    End If
                End If
            End If
        End With
    End If
    
    FreeChars PathChars
    
    GetRootLength = Result
End Function

Private Function IsSeparator(ByVal Ch As Integer) As Boolean
    IsSeparator = Ch = corDirectorySeparatorChar Or _
                  Ch = corAltDirectorySeparatorChar Or _
                  Ch = corVolumeSeparatorChar
End Function

Private Sub VerifyUNCPathFormat(ByVal NewBuffer As PathHelper, ByVal LengthToCheck As Long)
    Dim pos As Long
    pos = 2
    
    Do While pos < LengthToCheck
        If NewBuffer.Chars(pos) = corDirectorySeparatorChar Then
            pos = pos + 1
            Exit Do
        End If
        pos = pos + 1
    Loop

    If pos = LengthToCheck Then _
        Throw Error.Argument(Arg_PathIllegalUNC)
End Sub

Private Function GetExtensionCharIndex(ByRef Path As String) As Long
    Dim PathChars() As Integer
    PathChars = AllocChars(Path)
    
    Dim Index As Long
    Index = Len(Path) - 1
    
    Do While Index >= 0
        Select Case PathChars(Index)
            Case corPeriodChar
                GetExtensionCharIndex = Index + 1
                Exit Do
            Case corAltDirectorySeparatorChar, corDirectorySeparatorChar, corVolumeSeparatorChar
                Exit Do
        End Select
        Index = Index - 1
    Loop
    
    FreeChars PathChars
End Function

Private Function GetDirectoryCharIndex(ByRef Path As String) As Long
    Dim PathChars() As Integer
    PathChars = AllocChars(Path)
    
    Dim Index As Long
    Index = Len(Path) - 1
    
    Do While Index >= 0
        If IsSeparator(PathChars(Index)) Then
            GetDirectoryCharIndex = Index + 1
            Exit Do
        End If
        
        Index = Index - 1
    Loop
        
    FreeChars PathChars
End Function

Private Function FirstTwoChars(ByRef s As String) As DWord
    If LenB(s) > 0 Then
        FirstTwoChars = MemDWord(StrPtr(s))
    End If
End Function



