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

Private Const ParamBase                 As Long = 2000

Public Enum ErrorMessage
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
    Format_InvalidTimeSpan = 603
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
    Argument_LongerThanSrcArray = 821
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
    Overflow_TimeSpan = 1300
End Enum

Public Enum ParameterName
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
End Enum

Public Enum Param
    None = 0
    Index = 2000
    Count = 2001
    StartIndex = 2002
    Chars = 2003
    CharIndex = 2004
    CharCount = 2005
    ByteIndex = 2006
    Bytes = 2007
    ByteCount = 2008
    Value = 2009
    Arr = 2010
    List = 2011
    Year = 2012
    Month = 2013
    LCID = 2014
    Time = 2015
    PathParam = 2016
    DstArray = 2017
    StreamParam = 2018
    BufferParam = 2019
    Output = 2020
End Enum

Public Enum Rank
    MultiDimensionNotSupported = 200
End Enum

Public Enum ArgumentOutOfRange
    None = 0
    MustBeNonNegNum = 700
    SmallCapacity = 701
    NeedNonNegNum = 702
    ArrayListInsert = 703
    Index = 704
    LargerThanCollection = 705
    LowerBound = 706
    Exception = 707
    Range = 708
    UpperBound = 709
    MinMax = 710
    VersionFieldCount = 711
    ValidValues = 712
    NeedPosNum = 713
    OutsideConsoleBoundry = 714
    EnumType = 715
    ArrayLB = 716
    ArrayBounds = 717
    Count = 718
    NegativeLength = 719
    StartIndex = 720
    OffsetOut = 721
    IndexLength = 722
    InvalidFileTime = 723
    Month = 724
    Year = 725
    BeepFrequency = 726
    ConsoleBufferSize = 727
    ConsoleWindowSize_Size = 728
    ConsoleWindowPos = 729
    ConsoleBufferLessThanWindowSize = 730
    ConsoleTitleTooLong = 731
    ConsoleColor = 732
    CursorSize = 733
End Enum

Public Enum ArgumentNull
    None = 0
    NullArray = 900
    NullException = 901
    NullStream = 902
    NullCollection = 903
    NullTimeSpan = 904
    NullGeneric = 905
    NullBuffer = 906
End Enum

Public Enum Argument
    None = 0
    MultiDimensionNotSupported = 200
    InvalidOffLen = 800
    ArrayPlusOffTooSmall = 801
    Exception = 802
    ArrayRequired = 803
    MatchingBounds = 804
    IndexPlusTypeSize = 805
    VersionRequired = 806
    TimeSpanRequired = 807
    DateRequired = 808
    InvalidHandle = 809
    EmptyPath = 810
    SmallConversionBuffer = 811
    EmptyFileName = 812
    ReadableStreamRequired = 813
    InvalidEraValue = 814
    ParamRequired = 815
    StreamRequired = 816
    InvalidPathFormat = 817
    StreamNotReadable = 818
    StreamNotWritable = 819
    StreamNotSeekable = 820
    InvalidComparer = 823
    EmptyName = 825
    InvalidSeekOrigin = 827
    UnsupportedArray = 828
    NeedIntrinsicType = 829
    CharArrayRequired = 830
    InvalidDateSubtraction = 831
    InvalidPathChars = 832
End Enum

Public Enum ArgumentFormat
    LongerThanSrcArray = 821
    LongerThanDestArray = 822
    MustBeVbVarType = 824
    LongerThanSrcString = 826
End Enum


Public Function GetString(ByVal ResourceId As ErrorMessage, ParamArray Args() As Variant) As String
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(ModuleEBPOffset(4))
    GetString = cString.FormatArray(LoadResString(ResourceId), vArgs)
End Function

Public Function GetParameter(ByVal ParameterId As Param) As String
    If ParameterId <> Param.None Then
        GetParameter = LoadResString(ParameterId)
    End If
End Function

Public Function GetParameterName(ByVal Name As ParameterName)
    If Name <> Parameter_None Then
        GetParameterName = LoadResString(Name)
    End If
End Function

Public Function GetErrorMessage(ByVal Message As ErrorMessage) As String
    GetErrorMessage = LoadResString(Message)
End Function

Public Function GetErrorMessageFormat(ByVal Message As ErrorMessage, ParamArray Args() As Variant) As String
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(ModuleEBPOffset(4))
    GetErrorMessageFormat = cString.FormatArray(LoadResString(Message), SwappedArgs)
End Function

Private Function ModuleEBPOffset(ByVal Offset As Long) As Long
    On Error GoTo InIDE
    Debug.Assert 1 \ 0
    ModuleEBPOffset = Offset + 8
    Exit Function
    
InIDE:
    ModuleEBPOffset = Offset + 12
End Function
