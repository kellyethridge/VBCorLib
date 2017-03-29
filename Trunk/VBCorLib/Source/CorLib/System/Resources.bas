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

Public Enum ResourceStringKey
    Exception_WasThrown = 101
    ArrayTypeMismatch_Incompatible = 102
    ArrayTypeMismatch_Compare = 104
    Rank_MultiDimNotSupported = 105
    Rank_MustMatch = 106
    XMLSyntax_InvalidSyntax = 109

    UnauthorizedAccess_RegistryNoWrite = 150
    UnauthorizedAccess_IODenied_NoPathName = 151
    UnauthorizedAccess_IODenied_Path = 152
    
    Arg_PathIllegal = 200
    Arg_PathIllegalUNC = 201
    Arg_PathGlobalRoot = 202
    Arg_PathIsAVolume = 203
    Arg_ArrayPlusOffTooSmall = 204
    Arg_LongerThanSrcArray = 205
    Arg_LongerThanDestArray = 206
    Arg_LongerThanSrcString = 207
    Arg_BitArrayTypeUnsupported = 208
    Arg_ArrayLengthsDiffer = 209
    Arg_ArgumentException = 210
    Arg_ApplicationException = 211
    Arg_ArithmeticException = 212
    Arg_ArrayTypeMismatchException = 213
    Arg_CryptographyException = 214
    Arg_DirectoryNotFoundException = 215
    Arg_IOException = 216
    Arg_InvalidCastException = 217
    Arg_IndexOutOfRangeException = 218
    Arg_RankException = 219
    Arg_DriveNotFoundException = 220
    Arg_InvalidOperationException = 221
    Arg_OutOfMemoryException = 222
    Arg_FormatException = 223
    Arg_NotSupportedException = 224
    Arg_SerializationException = 225
    Arg_PlatformNotSupported = 226
    Arg_EndOfStreamException = 227
    Arg_OverflowException = 228
    Arg_UnauthorizedAccessException = 229
    Arg_ExternalException = 230
    Arg_ArgumentOutOfRangeException = 231
    Arg_SystemException = 232
    Arg_MustBeOfType = 233
    Arg_RankMultiDimNotSupported = 234
    Arg_FileIsDirectory_Name = 235
    Arg_VersionString = 236
    Arg_MustBeGuid = 237
    Arg_GuidArrayCtor = 238
    Arg_RegSubKeyAbsent = 239
    Arg_RegValStrLenBug = 240
    Arg_RegKeyStrLenBug = 241
    Arg_RegSubKeyValueAbsent = 242
    Arg_RegInvalidKeyName = 243
    
    Argument_MultiDimNotSupported = 105
    Argument_InvalidOffLen = 800
    Argument_ArrayRequired = 803
    Argument_MatchingBounds = 804 '?
    Argument_IndexPlusTypeSize = 805 '?
    Argument_VersionRequired = 806
    Argument_TimeSpanRequired = 807
    Argument_DateRequired = 808
    Argument_InvalidHandle = 809
    Argument_EmptyPath = 810
    Argument_SmallConversionBuffer = 811 '?
    Argument_EmptyFileName = 812
    Argument_ReadableStreamRequired = 813 '?
    Argument_InvalidEraValue = 814 '?
    Argument_ParamRequired = 815
    Argument_StreamRequired = 816
    Argument_InvalidPathFormat = 817 '?
    Argument_StreamNotReadable = 818
    Argument_StreamNotWritable = 819
    Argument_StreamNotSeekable = 820 '?
    Argument_InvalidComparer = 823 '?
    Argument_MustBeVbVarType = 824
    Argument_EmptyName = 825
    Argument_InvalidSeekOrigin = 827
    Argument_UnsupportedArray = 828
    Argument_NeedIntrinsicType = 829
    Argument_CharArrayRequired = 830
    Argument_InvalidDateSubtraction = 831 '?
    Argument_InvalidPathChars = 832
    Argument_InvalidValueType = 833 '?
    Argument_InvalidFileModeAndAccessCombo = 834 '?
    Argument_InvalidSeekOffset = 835 '?
    Argument_InvalidStreamSource = 836 '?
    Argument_NotEnumerable = 837 '?
    Argument_PathUriFormatNotSupported = 839
    Argument_ImplementIComparable = 840
    Argument_PathFormatNotSupported = 844
    Argument_AddingDuplicate = 845
    Argument_AddingDuplicate_Key = 846
    Argument_InvalidArrayLength = 847
    Argument_ByteArrayOrStreamRequired = 848
    Argument_InvalidValue = 849
    Argument_MinMaxValue = 850
    Argument_ByteArrayRequired = 851
    Argument_ByteArrayOrStringRequired = 852
    Argument_ByteArrayOrNumberRequired = 853
    Argument_InvalidElementTag = 854
    Argument_InvalidElementName = 855
    Argument_InvalidElementValue = 856
    Argument_InvalidElementText = 857
    Argument_AttributeNamesMustBeUnique = 858
    Argument_InvalidCharSequenceNoIndex = 859
    Argument_RecursiveFallbackBytes = 860
    Argument_InvalidCodePageBytesIndex = 861
    Argument_EncodingConversionOverflowChars = 862
    Argument_RecursiveFallback = 863
    Argument_InvalidHighSurrogate = 864
    Argument_InvalidLowSurrogate = 865
    Argument_InvalidCodePageConversionIndex = 866
    
    ArgumentNull_Array = 900
    ArgumentNull_Buffer = 901
    ArgumentNull_Stream = 902
    ArgumentNull_Collection = 903
    ArgumentNull_TimeSpan = 904
    ArgumentNull_Generic = 905
    ArgumentNull_Child = 906
    
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
    ArgumentOutOfRange_Bounds_Lower_Upper = 736
    ArgumentOutOfRange_DecimalScale = 737
    ArgumentOutOfRange_GetCharCountOverflow = 738
    ArgumentOutOfRange_GetByteCountOverflow = 739
    ArgumentOutOfRange_InvalidHighSurrogate = 740
    ArgumentOutOfRange_InvalidLowSurrogate = 741
    
    IOException_Exception = 400 '?
    IO_AlreadyExists_Name = 401 '?
    IO_FileTooLong2GB = 402
    IO_PathTooLong = 403
    IO_FileNotFound = 404
    IO_FileNotFound_Name = 405
    IO_PathNotFound_NoPathName = 406
    IO_PathNotFound_Path = 407
    IO_DriveNotFound_Drive = 408
    IO_SharingViolation_NoFileName = 409
    IO_SharingViolation_File = 410
    IO_FileExists_Name = 411
    
    Cryptography_HashNotYetFinalized = 500
    Cryptography_CSP_CFBSizeNotSupported = 501
    Cryptography_InvalidKeySize = 502
    Cryptography_InvalidIVSize = 503
    Cryptography_InvalidBlockSize = 504
    Cryptography_InvalidFeedbackSize = 505
    Cryptography_CSP_AlgorithmNotAvailable = 506
    Cryptography_PasswordDerivedBytes_FewBytesSalt = 507
    Cryptography_RC2_EKSKS2 = 508
    Cryptography_CryptoStream_FlushFinalBlockTwice = 509
    Cryptography_AddNullOrEmptyName = 510
    Cryptography_InvalidOID = 511

    Format_InvalidBase64Character = 600
    Format_InvalidNumberOfCharacters = 601
    Format_InvalidString = 602
    Format_InvalidTimeSpan = 603
    Format_InvalidGuidFormatSpecification = 604
    Format_GuidUnrecognized = 605
    Format_UnrecognizedEscapeSequence = 606
    
    IndexOutOfRange_Dimension = 300
    IndexOutOfRange_ArrayBounds = 301
    
    InvalidCast_FromTo = 1400
    
    InvalidOperation_EmptyStack = 1100
    InvalidOperation_EnumNotStarted = 1101
    InvalidOperation_EnumFinished = 1102
    InvalidOperation_VersionError = 1103
    InvalidOperation_EmptyQueue = 1104
    InvalidOperation_Comparer_Arg = 1105
    InvalidOperation_ReadOnly = 1106
    InvalidOperation_Timeouts = 1107
    InvalidOperation_WrongAsyncResultOrEndReadCalledMultiple = 1108
    InvalidOperation_WrongAsyncResultOrEndWriteCalledMultiple = 1109
    InvalidOperation_RegRemoveSubKey = 1110
    
    NotSupported_ReadOnlyCollection = 1000
    NotSupported_FixedSizeCollection = 1001
    NotSupported_MemoryStreamNotExpandable = 1002
    NotSupported_UnwritableStream = 1003
    NotSupported_UnreadableStream = 1004
    NotSupported_UnseekableStream = 1005
    NotSupported_FileStreamOnNonFiles = 1006
    
    ObjectDisposed_StreamClosed = 1200
    ObjectDisposed_FileNotOpen = 1201
    ObjectDisposed_Generic = 1202
    ObjectDisposed_ReaderClosed = 1203
    ObjectDisposed_WriterClosed = 1204
    ObjectDisposed_RegKeyClosed = 1205
    
    Overflow_TimeSpan = 1300
    
    UnknownError_Num = 1500
End Enum

Public Enum ParameterResourceKey
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
    Parameter_Values = 2027
    Parameter_InArray = 2028
    Parameter_OutArray = 2029
    Parameter_Offset = 2030
    Parameter_InputBuffer = 2031
    Parameter_InputOffset = 2032
    Parameter_InputCount = 2033
    Parameter_OutputBuffer = 2034
    Parameter_OutputOffset = 2035
    Parameter_RgbHash = 2036
    Parameter_Signature = 2037
End Enum
