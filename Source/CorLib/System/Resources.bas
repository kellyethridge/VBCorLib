Attribute VB_Name = "Resources"
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
' Module: Resources
'
Option Explicit

Public Enum ResourceStringId
    Exception_WasThrown = 101
    ArrayTypeMismatch_Incompatible = 102
    ArrayTypeMismatch_Exception = 103
    ArrayTypeMismatch_Compare = 104
    Rank_MultiDimension = 200
    IndexOutOfRange_Dimension = 300
    IOException_Exception = 400
    IOException_DirectoryExists = 401
    FileNotFound_Exception = 500
    Format_InvalidBase64Character = 600
    Format_InvalidNumberOfCharacters = 601
    ArgumentOutOfRange_MustBeNonNegNum = 1000
    ArgumentOutOfRange_SmallCapacity = 1001
    ArgumentOutOfRange_NeedNonNegNum = 1002
    ArgumentOutOfRange_ArrayListInsert = 1003
    ArgumentOutOfRange_Index = 1004
    ArgumentOutOfRange_LargerThanCollection = 1005
    ArgumentOutOfRange_LBound = 1006
    ArgumentOutOfRange_Exception = 1007
    ArgumentOutOfRange_Range = 1008
    ArgumentOutOfRange_UBound = 1009
    ArgumentOutOfRange_MinMax = 1010
    ArgumentOutOfRange_VersionFieldCount = 1011
    ArgumentOutOfRange_ValidValues = 1012
    ArgumentOutOfRange_NeedPosNum = 1013
    ArgumentOutOfRange_OutsideConsoleBoundry = 1014
    ArgumentOutOfRange_Enum = 1015
    Argument_InvalidCountOffset = 2000
    Argument_ArrayPlusOffTooSmall = 2001
    Argument_Exception = 2002
    Argument_ArrayRequired = 2003
    Argument_MatchingBounds = 2004
    Argument_IndexPlusTypeSize = 2005
    Argument_VersionRequired = 2006
    Argument_TimeSpanRequired = 2007
    Argument_DateRequired = 2008
    Argument_InvalidHandle = 2009
    Argument_EmptyPath = 2010
    Argument_SmallConversionBuffer = 2011
    Argument_EmptyFileName = 2012
    Argument_ReadableStreamRequired = 2013
    Argument_InvalidEraValue = 2014
    Argument_ParamRequired = 2015
    Argument_StreamRequired = 2016
    Argument_InvalidPathFormat = 2017
    Argument_StreamNotReadable = 2018
    Argument_StreamNotWritable = 2019
    Argument_StreamNotSeekable = 2020
    ArgumentNull_Array = 2100
    ArgumentNull_Exception = 2101
    ArgumentNull_Stream = 2102
    ArgumentNull_Collection = 2103
    ArgumentNull_TimeSpan = 2104
    NotSupported_ReadOnlyCollection = 3000
    NotSupported_FixedSizeCollection = 3001
    InvalidOperation_EmptyStack = 4000
    InvalidOperation_EnumNotStarted = 4001
    InvalidOperation_EnumFinished = 4002
    InvalidOperation_VersionError = 4003
    InvalidOperation_EmptyQueue = 4004
    InvalidOperation_Comparer_Arg = 4005
    InvalidOperation_ReadOnly = 4006
    InvalidOperation_Timeouts = 4007
End Enum

Private mBuilder As New StringBuilder

Public Function GetString(ByVal ResourceId As ResourceStringId, ParamArray Args() As Variant) As String
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(ModuleEBPOffset(4)))
    GetString = cString.FormatArray(LoadResString(ResourceId), vArgs)
End Function

Public Function ModuleEBPOffset(ByVal Offset As Long) As Long
    On Error GoTo InIDE
    Debug.Assert 1 \ 0
    ModuleEBPOffset = Offset + 8
    Exit Function
    
InIDE:
    ModuleEBPOffset = Offset + 12
End Function
