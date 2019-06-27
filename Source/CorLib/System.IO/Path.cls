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
Public Property Get AltDirectorySeparator() As String
    AltDirectorySeparator = vbAltDirectorySeparator
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
Public Property Get DirectorySeparator() As String
    DirectorySeparator = vbDirectorySeparator
End Property

''
' The character code used to separate strings in an environment variable.
'
' @return The environment variable string separator character code (59).
'
Public Property Get PathSeparatorChar() As Integer
    PathSeparatorChar = vbPathSeparatorChar
End Property

''
' The string used to seperate strings in an environment variable.
'
' @return The environment variable string separator ";".
'
Public Property Get PathSeparator() As String
    PathSeparator = vbPathSeparator
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
Public Property Get VolumeSeparator() As String
    VolumeSeparator = vbVolumeSeparator
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
        ElseIf AscW(Extension) = vbPeriodChar Then
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
    If Len(Path) = 0 Then
        Exit Function
    End If
    
    CheckValidPathChars Path
    
    With FirstTwoChars(Path)
        If IsDirectorySeparator(.LoWord) Then
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
Public Function HasExtension(ByRef Path As String) As Boolean
    Dim Index As Long
    
    CheckValidPathChars Path
    Index = GetExtensionCharIndex(Path)
    HasExtension = Index > 0 And Index < Len(Path)
End Function

''
' Combintes two paths into a single path.
'
' @param Path1 The left-hand side of the two paths to be joined.
' @param Path2 The right-hand side of the two paths to be joined.
' @param Paths Additional paths to be joined.
' @return The joined paths.
' @remarks <p>If Path1 is an empty string, then Path2 is returned as is. If Path2
' is empty, but not Path1, then Path1 is returned as is.</p>
' <p>If Path2 is rooted (starts with a volume, UNC identifier or directory
' separator), then Path2 is returned, regardless of Path1.</p>
' <p>If Path1 does not end with a directory separator, one will be inserted.</p>
' <p>This function does not verify that the paths exist.</p>
'
Public Function Combine(ByRef Path1 As String, ByRef Path2 As String, ParamArray Paths() As Variant) As String
    Dim Result  As String
    Dim Part    As Variant
    
    CheckValidPathChars Path1
    Result = CombinePaths(Path1, Path2)
    
    For Each Part In Paths
        Result = CombinePaths(Result, Object.ToString(Part))
    Next
    
    Combine = Result
End Function

''
' Combines an array of path fragments returning a complete path.
'
' @param Paths The array of path fragments to be combined.
' @return Returns a concatenated string of the path fragments.
Public Function CombineArray(ByRef Paths() As String) As String
    Dim Index   As Long
    Dim Result  As String
    
    If CorArray.IsNullOrEmpty(Paths) Then
        Exit Function
    End If
    
    Index = LBound(Paths)
    Result = Paths(Index)
    CheckValidPathChars Result
    
    For Index = Index + 1 To UBound(Paths)
        Result = CombinePaths(Result, Paths(Index))
    Next
    
    CombineArray = Result
End Function

''
' Returns the directory portion of a file path.
'
' @param Path The path to parse the directory from.
' @return The directory parsed from the path.
' @remarks The function does not verify that the directory exists.
'
Public Function GetDirectoryName(ByRef Path As String) As String
    Dim NormalizedPath  As String
    Dim PathLength      As Long
    Dim RootLength      As Long
    
    If Len(Path) = 0 Then _
        Error.Argument Argument_EmptyPath, "Path"
    
    CheckValidPathChars Path
    NormalizedPath = NormalizePath(Path, False)
    RootLength = GetRootLength(NormalizedPath)
    
    If RootLength < Len(NormalizedPath) Then
        PathLength = InStrRev(NormalizedPath, vbDirectorySeparator)
        
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
    Dim Index As Long
    
    CheckValidPathChars Path
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
    Dim Index    As Long
    
    FileName = GetFileName(Path)
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
    
    If InStr(3, FullPath, vbVolumeSeparator) >= 3 Then _
        Error.NotSupported Argument_PathFormatNotSupported
        
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
    
    FileName = String$(MAX_PATH, 0)
    
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
    
    PathName = String$(MAX_PATH, 0)
    Size = GetTempPathW(MAX_PATH, PathName)
    
    If Size = 0 Or Size > MAX_PATH Then
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mInvalidPathChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F)
    mInvalidFileNameChars = Cor.NewIntegers(&H22, &H3C, &H3E, &H7C, &H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF, &H10, &H11, &H12, &H13, &H14, &H15, &H16, &H17, &H18, &H19, &H1A, &H1B, &H1C, &H1D, &H1E, &H1F, &H3A, &H2A, &H3F, &H5C, &H2F)
    mTrimEndChars = Cor.NewIntegers(&H9, &HA, &HB, &HC, &HD, &H20, &H85, &HA0)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRootLength(ByRef Path As String) As Long
    Dim PathChars() As Integer
    Dim Result      As Long
    Dim Length      As Long
    
    PathChars = AllocChars(Path)
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
            ElseIf .HiWord = vbVolumeSeparatorChar Then
                Result = 2
                
                ' check for following directory separator
                If Length > 2 Then
                    If PathChars(2) = vbDirectorySeparatorChar Then
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
    IsSeparator = Ch = vbDirectorySeparatorChar Or _
                  Ch = vbAltDirectorySeparatorChar Or _
                  Ch = vbVolumeSeparatorChar
End Function

Private Sub VerifyUNCPathFormat(ByVal NewBuffer As PathHelper, ByVal LengthToCheck As Long)
    Dim pos As Long
    
    pos = 2
    
    Do While pos < LengthToCheck
        If NewBuffer.Chars(pos) = vbDirectorySeparatorChar Then
            pos = pos + 1
            Exit Do
        End If
        
        pos = pos + 1
    Loop

    If pos = LengthToCheck Then _
        Error.Argument Arg_PathIllegalUNC
End Sub

Private Function GetExtensionCharIndex(ByRef Path As String) As Long
    Dim PathChars() As Integer
    Dim Index       As Long
    
    PathChars = AllocChars(Path)
    Index = Len(Path) - 1
    
    Do While Index >= 0
        Select Case PathChars(Index)
            Case vbPeriodChar
                GetExtensionCharIndex = Index + 1
                Exit Do
            Case vbAltDirectorySeparatorChar, vbDirectorySeparatorChar, vbVolumeSeparatorChar
                Exit Do
        End Select
        
        Index = Index - 1
    Loop
    
    FreeChars PathChars
End Function

Private Function GetDirectoryCharIndex(ByRef Path As String) As Long
    Dim PathChars() As Integer
    Dim Index       As Long
    
    PathChars = AllocChars(Path)
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

Private Function CombinePaths(ByRef Path1 As String, ByRef Path2 As String) As String
    If IsPathRooted(Path2) Then
        CombinePaths = Path2
        Exit Function
    End If
    
    If Len(Path1) = 0 Then
        CombinePaths = Path2
        Exit Function
    ElseIf Len(Path2) = 0 Then
        CombinePaths = Path1
        Exit Function
    End If
    
    If IsSeparator(CorString.LastChar(Path1)) Then
        CombinePaths = Path1 & Path2
    Else
        CombinePaths = Path1 & vbDirectorySeparator & Path2
    End If
End Function

Private Sub CheckValidPathChars(ByRef Path As String)
    If HasIllegalCharacters(Path) Then
        Error.Argument Argument_InvalidPathChars
    End If
End Sub

Private Function HasIllegalCharacters(ByRef Path As String) As Boolean
    Dim PathChars() As Integer
    Dim i           As Long
    
    PathChars = AllocChars(Path)
    
    For i = 0 To Len(Path) - 1
        Select Case PathChars(i)
            Case vbQuoteChar, vbLessThanChar, vbGreaterThanChar, vbPipeChar, Is < vbSpaceChar
                HasIllegalCharacters = True
                Exit For
        End Select
    Next
    
    FreeChars PathChars
End Function

Private Function NormalizePath(ByRef Path As String, ByVal FullCheck As Boolean) As String
    Dim NewBuffer   As New PathHelper
    Dim PathChars() As Integer
    Dim Index       As Long
    Dim MaxIndex    As Long
    
    If FullCheck Then
        PathChars = AllocChars(CleanPath(Path))
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
        Error.Argument Arg_PathIllegal
    
    ValidateUNCPath NewBuffer, FullCheck
    NormalizePath = NewBuffer.ToString
End Function

Private Function CleanPath(ByRef Path As String) As String
    CleanPath = CorString.TrimEnd(Path, mTrimEndChars)
    
    If HasIllegalCharacters(CleanPath) Then
        Error.Argument Argument_InvalidPathChars
    End If
End Function

Private Sub ValidateUNCPath(ByRef NewBuffer As PathHelper, ByVal FullCheck As Boolean)
    Dim MaxUNCPathCheck As Long
    
    If FullCheck Then
        If NewBuffer.StartsWith("http:") Or NewBuffer.StartsWith("file:") Then _
            Error.Argument Argument_PathUriFormatNotSupported
    
        MaxUNCPathCheck = NewBuffer.GetFullPathName
    Else
        MaxUNCPathCheck = 1
    End If
    
    If MaxUNCPathCheck <> 0 Then
        If NewBuffer.IsUNCPath Then
            VerifyUNCPathFormat NewBuffer, MaxUNCPathCheck
        End If
    End If
End Sub

Private Function IsDirectorySeparator(ByVal Ch As Integer) As Boolean
    IsDirectorySeparator = Ch = vbDirectorySeparatorChar Or _
                           Ch = vbAltDirectorySeparatorChar
End Function

Private Function PrepareForUNCPath(ByRef PathChars() As Integer, ByRef Index As Long, ByVal NewBuffer As PathHelper) As Boolean
    If UBound(PathChars) > 0 Then
        If IsDirectorySeparator(PathChars(0)) Then
            NewBuffer.Append vbDirectorySeparatorChar
            Index = 1
            PrepareForUNCPath = True
        End If
    End If
End Function

Private Sub PrepareForVolumeLabel(ByRef PathChars() As Integer, ByRef Index As Long, ByVal NewBuffer As PathHelper)
    If UBound(PathChars) > 0 Then
        If PathChars(1) = vbVolumeSeparatorChar Then
            Select Case PathChars(0)
                Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar
                    NewBuffer.SetVolumeLabel PathChars(0)
                    Index = 2
                Case Else
                    Error.Argument Arg_PathIllegal
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
            Case vbPeriodChar
                HasPeriods = True
                LastPeriodIndex = pos
            Case vbSpaceChar
                HasSpaces = True
                LastSpaceIndex = pos
            Case vbDirectorySeparatorChar, vbAltDirectorySeparatorChar
                If pos > Index Then
                    Exit Do
                End If
                
                If Not FoundSlash Then
                    NewBuffer.Append vbDirectorySeparatorChar
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
            Error.Argument Arg_PathIllegal
        End If
    
        LastCharIndex = CorMath.Min(Index + 1, LastPeriodIndex)
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
            Case vbTildaChar
                TryExpand = True
            Case vbDirectorySeparatorChar, vbAltDirectorySeparatorChar
                Exit Do
            Case vbPeriodChar, vbSpaceChar
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



