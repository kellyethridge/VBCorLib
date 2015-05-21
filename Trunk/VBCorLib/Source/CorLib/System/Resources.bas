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

Public Enum ResourceString
    None = 0
    Exception_WasThrown = 101
    ArrayTypeMismatch_Incompatible = 102
    ArrayTypeMismatch_Exception = 103
    ArrayTypeMismatch_Compare = 104
    Rank_MultiDimNotSupported = 200
    IOException_Exception = 400
    IOException_DirectoryExists = 401
    IOException_FileTooLong2GB = 402
    FileNotFound_Exception = 500
    Format_InvalidBase64Character = 600
    Format_InvalidNumberOfCharacters = 601
    Format_InvalidString = 602
    Format_InvalidTimeSpan = 603
    Overflow_TimeSpan = 1300
    InvalidCast_FromTo = 1400
End Enum

Public Enum IOExceptionString
    IOException_Exception = 400
    IOException_DirectoryExists = 401
    IOException_FileTooLong2GB = 402
    IOException_PathTooLong = 403
End Enum

Public Enum NotSupportedString
    NotSupported_ReadOnlyCollection = 1000
    NotSupported_FixedSizeCollection = 1001
    NotSupported_MemoryStreamNotExpandable = 1002
    NotSupported_UnwritableStream = 1003
    NotSupported_UnreadableStream = 1004
    NotSupported_UnseekableStream = 1005
End Enum

Public Enum InvalidOperationString
    InvalidOperation_EmptyStack = 1100
    InvalidOperation_EnumNotStarted = 1101
    InvalidOperation_EnumFinished = 1102
    InvalidOperation_VersionError = 1103
    InvalidOperation_EmptyQueue = 1104
    InvalidOperation_Comparer_Arg = 1105
    InvalidOperation_ReadOnly = 1106
    InvalidOperation_Timeouts = 1107
    InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple = 1108
End Enum

Public Enum IndexOutOfRangeString
    IndexOutOfRange_Dimension = 300
    IndexOutOfRange_ArrayBounds = 301
End Enum

Public Enum ObjectDisposedString
    ObjectDisposed_StreamClosed = 1200
    ObjectDisposed_FileNotOpen = 1201
    ObjectDisposed_Generic = 1202
End Enum

Public Enum ArgumentString
    Argument_MultiDimNotSupported = ResourceString.Rank_MultiDimNotSupported
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
    Argument_LongerThanSrCorArray = 821
    Argument_LongerThanDestArray = 822
    Argument_InvalidComparer = 823
    Argument_MustBeVbVarType = 824
    Argument_EmptyName = 825
    Argument_LongerThanSrcString = 826
    Argument_InvalidSeekOrigin = 827
    Argument_UnsupportedArray = 828
    Argument_NeedIntrinsicType = 829
    Argument_CharArrayRequired = 830
    Argument_InvalidDateSubtraction = 831
    Argument_InvalidCharsInPath = 832
    Argument_InvalidValueType = 833
    Argument_InvalidFileModeAndAccessCombo = 834
    Argument_InvalidSeekOffset = 835
    Argument_InvalidStreamSource = 836
    Argument_NotEnumerable = 837
    Arg_PathIllegal = 838
    Argument_PathUriFormatNotSupported = 839
    Arg_PathIllegalUNC = 840
    Arg_PathGlobalRoot = 841
    Argument_BitArrayTypeUnsupported = 842
End Enum

Public Enum ArgumentNullString
    ArgumentNull_Array = 900
    ArgumentNull_Exception = 901
    ArgumentNull_Stream = 902
    ArgumentNull_Collection = 903
    ArgumentNull_TimeSpan = 904
    ArgumentNull_Generic = 905
    ArgumentNull_Buffer = 906
End Enum

Public Enum ArgumentOutOfRangeString
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
    ArgumentOutOfRange_IndexCountBuffer = 734
    ArgumentOutOfRange_IndexCount = 735
End Enum

Public Enum ParameterString
    Parameter_None = 0
    Parameter_Index = 2000
    Parameter_Count = 2001
    Parameter_StartIndex = 2002
    Parameter_Chars = 2003
    Parameter_CharIndex = 2004
    Parameter_CharCount = 2005
    Parameter_ByteIndex = 2006
    Parameter_Bytes = 2007
    Parameter_ByteCount = 2008
    Parameter_Value = 2009
    Parameter_Arr = 2010
    Parameter_List = 2011
    Parameter_Year = 2012
    Parameter_Month = 2013
    Parameter_LCID = 2014
    Parameter_Time = 2015
    Parameter_Path = 2016
    Parameter_DstArray = 2017
    Parameter_Stream = 2018
    Parameter_Buffer = 2019
    Parameter_Output = 2020
    Parameter_Source = 2021
    Parameter_Length = 2022
    Parameter_SourceArray = 2023
    Parameter_DestinationArray = 2024
    Parameter_Keys = 2025
    Parameter_Items = 2026
End Enum
