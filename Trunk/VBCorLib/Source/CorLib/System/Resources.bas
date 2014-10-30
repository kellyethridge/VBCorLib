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
    Rank_MultiDimNotSupported = 200
    IndexOutOfRange_Dimension = 300
    IOException_Exception = 400
    IOException_DirectoryExists = 401
    FileNotFound_Exception = 500
    Format_InvalidBase64Character = 600
    Format_InvalidNumberOfCharacters = 601
    Format_InvalidString = 602
    ArgumentOutOfRange_MustBeNonNegNum = 700
    ArgumentOutOfRange_SmallCapacity = 701
    ArgumentOutOfRange_NeedNonNegNum = 702
    ArgumentOutOfRange_ArrayListInsert = 703
    ArgumentOutOfRange_Index = 704
    ArgumentOutOfRange_LargerThanCollection = 705
    ArgumentOutOfRange_LBound = 706
    ArgumentOutOfRange_Exception = 707
    ArgumentOutOfRange_Range = 708
    ArgumentOutOfRange_UBound = 709
    ArgumentOutOfRange_MinMax = 710
    ArgumentOutOfRange_VersionFieldCount = 711
    ArgumentOutOfRange_ValidValues = 712
    ArgumentOutOfRange_NeedPosNum = 713
    ArgumentOutOfRange_OutsideConsoleBoundry = 714
    ArgumentOutOfRange_Enum = 715
    ArgumentOutOfRange_ArrayLB = 716
    ArgumentOutOfRange_ArrayBounds = 717
    ArgumentOutOfRange_Count = 718
    ArgumentOutOfRange_NegativeLength = 719
    ArgumentOutOfRange_StartIndex = 720
    ArgumentOutOfRange_OffsetOut = 721
    ArgumentOutOfRange_IndexLength = 722
    ArgumentOutOfRange_InvalidFileTime = 723
    ArgumentOutOfRange_Month = 724
    ArgumentOutOfRange_Year = 725
    ArgumentOutOfRange_BeepFrequency = 726
    ArgumentOutOfRange_ConsoleBufferSize = 727
    ArgumentOutOfRange_ConsoleWindowSize_Size = 728
    ArgumentOutOfRange_ConsoleWindowPos = 729
    ArgumentOutOfRange_ConsoleBufferLessThanWindowSize = 730
    ArgumentOutOfRange_ConsoleTitleTooLong = 731
    ArgumentOutOfRange_ConsoleColor = 732
    ArgumentOutOfRange_CursorSize = 733
    Argument_InvalidOffLen = 800
    Argument_ArrayPlusOffTooSmall = 801
    Argument_Exception = 802
    Argument_ArrayRequired = 803
    Argument_MatchingBounds = 804
    Argument_IndexPlusTypeSize = 805
    Argument_VersionRequired = 806
    Argument_TimeSpanRequired = 807
    Argument_DateRequired = 808
    Argument_InvalidHandle = 809
    Argument_EmptyPath = 810
    Argument_SmallConversionBuffer = 811
    Argument_EmptyFileName = 812
    Argument_ReadableStreamRequired = 813
    Argument_InvalidEraValue = 814
    Argument_ParamRequired = 815
    Argument_StreamRequired = 816
    Argument_InvalidPathFormat = 817
    Argument_StreamNotReadable = 818
    Argument_StreamNotWritable = 819
    Argument_StreamNotSeekable = 820
    Arg_LongerThanSrcArray = 821
    Arg_LongerThanDestArray = 822
    Argument_InvalidComparer = 823
    Arg_MustBeVbVarType = 824
    Argument_EmptyName = 825
    Arg_LongerThanSrcString = 826
    Argument_InvalidSeekOrigin = 827
    Argument_UnsupportedArray = 828
    Argument_NeedIntrinsicType = 829
    Argument_CharArrayRequired = 830
    ArgumentNull_Array = 900
    ArgumentNull_Exception = 901
    ArgumentNull_Stream = 902
    ArgumentNull_Collection = 903
    ArgumentNull_TimeSpan = 904
    ArgumentNull_Generic = 905
    ArgumentNull_Buffer = 906
    NotSupported_ReadOnlyCollection = 1000
    NotSupported_FixedSizeCollection = 1001
    NotSupported_MemoryStreamNotExpandable = 1002
    NotSupported_UnwritableStream = 1003
    InvalidOperation_EmptyStack = 1100
    InvalidOperation_EnumNotStarted = 1101
    InvalidOperation_EnumFinished = 1102
    InvalidOperation_VersionError = 1103
    InvalidOperation_EmptyQueue = 1104
    InvalidOperation_Comparer_Arg = 1105
    InvalidOperation_ReadOnly = 1106
    InvalidOperation_Timeouts = 1107
    ObjectDisposed_StreamClosed = 1200
End Enum

Public Enum ParameterResourceId
    Param_None = 0
    Param_Index = 2000
    Param_Count
    Param_StartIndex
    Param_Chars
    Param_CharIndex
    Param_CharCount
    Param_ByteIndex
    Param_Bytes
    Param_ByteCount
End Enum

Public Function GetString(ByVal ResourceId As ResourceStringId, ParamArray Args() As Variant) As String
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(ModuleEBPOffset(4))
    GetString = cString.FormatArray(LoadResString(ResourceId), vArgs)
End Function

Public Function GetParameter(ByVal ParameterId As ParameterResourceId) As String
    GetParameter = LoadResString(ParameterId)
End Function

Public Function ModuleEBPOffset(ByVal Offset As Long) As Long
    On Error GoTo InIDE
    Debug.Assert 1 \ 0
    ModuleEBPOffset = Offset + 8
    Exit Function
    
InIDE:
    ModuleEBPOffset = Offset + 12
End Function
