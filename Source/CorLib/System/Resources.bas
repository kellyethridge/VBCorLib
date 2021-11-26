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
    Arg_InvalidSearchPattern = 244
    Arg_MustBeDateTime = 245
    Arg_UnsupportedResourceType = 246
    Arg_TypeNotSupported = 247
    
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
    Argument_InvalidStreamSource = 836
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
    Argument_EncodingConversionOverflowBytes = 867
    Argument_FallbackBufferNotEmpty = 868
    Argument_ConversionOverflow = 869
    Argument_EncodingNotSupported = 870
    Argument_InvalidCodePageBytes = 871
    Argument_InvalidCodePageChars = 872
    Argument_IntegerRequired = 873
    Argument_ResultCalendarRange = 874
    Argument_InvalidLanguageIdSource = 875
    Argument_InvalidResourceNameOrType = 876
    Argument_MaxStringLength = 877
    Argument_InvalidResourceKeyType = 878
    Argument_StringZeroLength = 879
    Argument_EmptyIniSection = 880
    Argument_EmptyIniKey = 881
    Argument_MapNameEmptyString = 882
    Argument_NewMMFAppendModeNotAllowed = 883
    Argument_NewMMFWriteAccessNotAllowed = 884
    Argument_EmptyFile = 885
    Argument_NotEnoughBytesToRead = 886
    Argument_InvalidStructure = 887
    Argument_ReadAccessWithLargeCapacity = 888
    
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
    ArgumentOutOfRange_Day = 742
    ArgumentOutOfRange_CalendarRange = 743
    ArgumentOutOfRange_BadYearMonthDay = 744
    ArgumentOutOfRange_AddValue = 745
    ArgumentOutOfRange_Era = 746
    ArgumentOutOfRange_PositiveOrDefaultCapacityRequired = 747
    ArgumentOutOfRange_CapacityGEFileSizeRequired = 748
    ArgumentOutOfRange_PositiveOrDefaultSizeRequired = 749
    ArgumentOutOfRange_CapacityLargerThanLogicalAddressSpaceNotAllowed = 750
    ArgumentOutOfRange_PositionLessThanCapacityRequired = 751
    ArgumentOutOfRange_StreamLength = 752
    ArgumentOutOfRange_NegativeCount = 753
    ArgumentOutOfRange_OffsetLength = 754
    ArgumentOutOfRange_InvalidUTF32 = 755
    
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
    IO_SourceDestMustBeDifferent = 412
    IO_SourceDestMustHaveSameRoot = 413
    IO_EOF_ReadBeyondEOF = 414
    
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
    Cryptography_InvalidPaddingMode = 512
    Cryptography_InvalidCipherMode = 513
    Cryptography_InvalidKeyForState = 514
    Cryptography_InvalidDSASignatureSize = 515

    Format_InvalidBase64Character = 600
    Format_InvalidNumberOfCharacters = 601
    Format_InvalidString = 602
    Format_InvalidTimeSpan = 603
    Format_InvalidGuidFormatSpecification = 604
    Format_GuidUnrecognized = 605
    Format_UnrecognizedEscapeSequence = 606
    Format_ParseBigInteger = 607
    Format_BadQuote = 608
    Format_IndexOutOfRange = 609
    Format_BadFormatSpecifier = 610
    Format_NeedSingleChar = 611
    Format_Bad7BitInt32 = 612
    
    IndexOutOfRange_Dimension = 300
    IndexOutOfRange_ArrayBounds = 301
    
    InvalidCast_FromTo = 1400
    InvalidCast_DownCastArrayElement = 1401
    InvalidCast_IComparer = 1402
    InvalidCast_StoreArrayElement = 1403
    
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
    InvalidOperation_ResourceWriterSaved = 1111
    InvalidOperation_GetVersion = 1112
    InvalidOperation_ConsoleReadKeyOnFile = 1113
    
    NotSupported_ReadOnlyCollection = 1000
    NotSupported_FixedSizeCollection = 1001
    NotSupported_MemoryStreamNotExpandable = 1002
    NotSupported_UnwritableStream = 1003
    NotSupported_UnreadableStream = 1004
    NotSupported_UnseekableStream = 1005
    NotSupported_FileStreamOnNonFiles = 1006
    NotSupported_Reading = 1007
    NotSupported_Writing = 1008
    NotSupported_MMViewStreamsFixedLength = 1009
    NotSupported_StringComparison = 1010
    
    ObjectDisposed_StreamClosed = 1200
    ObjectDisposed_FileNotOpen = 1201
    ObjectDisposed_Generic = 1202
    ObjectDisposed_ReaderClosed = 1203
    ObjectDisposed_WriterClosed = 1204
    ObjectDisposed_RegKeyClosed = 1205
    
    Overflow_TimeSpan = 1300
    Overflow_Int64 = 1301
    Overflow_Char = 1302
    
    UnknownError_Num = 1500
    
    LastDllError = 10000
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

Public Enum ParameterName
    NameOfIndex = 2000
    NameOfCount = 2001
    NameOfStartIndex = 2002
    NameOfChars = 2003
    NameOfCharIndex = 2004
    NameOfCharCount = 2005
    NameOfByteIndex = 2006
    NameOfBytes = 2007
    NameOfByteCount = 2008
    NameOfValue = 2009
    NameOfArr = 2010
    NameOfList = 2011
    NameOfYear = 2012
    NameOfMonth = 2013
    NameOfLCID = 2014
    NameOfTime = 2015
    NameOfPath = 2016
    NameOfDstArray = 2017
    NameOfStream = 2018
    NameOfBuffer = 2019
    NameOfOutput = 2020
    NameOfSource = 2021
    NameOfLength = 2022
    NameOfSourceArray = 2023
    NameOfDestinationArray = 2024
    NameOfKeys = 2025
    NameOfItems = 2026
    NameOfValues = 2027
    NameOfInArray = 2028
    NameOfOutArray = 2029
    NameOfOffset = 2030
    NameOfInputBuffer = 2031
    NameOfInputOffset = 2032
    NameOfInputCount = 2033
    NameOfOutputBuffer = 2034
    NameOfOutputOffset = 2035
    NameOfRgbHash = 2036
    NameOfSignature = 2037
    NameOfRgb = 2038
    NameOfKeyBlob = 2039
    NameOfArrayIndex = 2040
    NameOfDestination = 2041
    NameOfa = 2042
    NameOfb = 2043
End Enum

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1 To 256) As Long
End Type

Public Function CursorResourceFromHandle(ByVal Handle As Long) As Byte()
    Dim Result() As Byte
    
    If SaveHICONtoArray(Handle, Result, False) Then
        CursorResourceFromHandle = Result
    End If
End Function

Public Function IconResourceFromHandle(ByVal Handle As Long) As Byte()
    Dim Result() As Byte
    
    If SaveHICONtoArray(Handle, Result, True) Then
        IconResourceFromHandle = Result
    End If
End Function

Public Sub ValidateResourceName(ByRef ResourceName As Variant)
    If Not IsValidResourceNameOrType(ResourceName) Then
        Error.Argument Argument_InvalidResourceNameOrType, "ResourceName"
    End If
End Sub

Public Sub ValidateResourceType(ByRef ResourceType As Variant)
    If Not IsValidResourceNameOrType(ResourceType) Then
        Error.Argument Argument_InvalidResourceNameOrType, "ResourceType"
    End If
End Sub

Private Function IsValidResourceNameOrType(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbString, vbLong, vbInteger, vbByte
            IsValidResourceNameOrType = True
    End Select
End Function

' using LaVolpe's method for converting an icon or cursor to a byte array vs getting
' an IPicture to save itself out because his method produces a much better result.
'
' http://www.vbforums.com/showthread.php?637452-vb6-Icon-Handle-to-File-Array
'
Private Function SaveHICONtoArray(ByVal hIcon As Long, OutArray() As Byte, ByVal IsIcon As Boolean) As Boolean

    ' Function takes an HICON handle and converts it to 1,4,8,24, or 32 bit icon file format
    ' If return value is False, outArray() contents are undefined
    ' Note: Bit reduction is in play. Example: If original source for HICON was 24 bit
    '   and it can be reduced/saved as 8 bit or lower without color loss, it will.
    ' Note: The end result's quality should be identical to HICON
    '       XP and above required to show/save 32bpp alphablended icons correctly
    '       This routine not coded to save icons in PNG format (Vista and above)

    Dim Bits() As Long, pow2(0 To 8) As Long
    Dim tDC As Long, maskScan As Long, clrScan As Long
    Dim x As Long, y As Long, clrOffset As Long, bNewColor As Boolean
    Dim palIndex As Long, palShift As Long, palPtr As Long, lPrevPal As Long
    Dim ICI As ICONINFO, BHI As BITMAPINFO
    
    If hIcon = 0& Then Exit Function
    If GetIconInfo(hIcon, ICI) = 0& Then Exit Function
    
    If IsIcon <> CBool(ICI.fIcon = BOOL_TRUE) Then
        Exit Function
    End If
    
    ' A properly formatted icon file will contain this information:
    ' :: 6 byte ICONDIRECTORY structure
    ' :: 16 byte ICONDIRECTORYENTRY structure
    ' If stored in PNG format then
    '    :: The entire PNG
    ' Else
    '    :: 40 byte BITMAPINFOHEADER structure
    '    If paletted then: Palette entries, each in BGRA format
    '    :: Color data packed & word-aligned per Bitcount of 1,4,8,24,32 bits per pixel
    '    :: 1-bit word-aligned Mask data, even if mask not used (i.e., 32bpp)
    '    Size of any single icon's file can be calculated as:
    '    FileSize = 62 + NrPaletteEntries*4 + (ByteAlignOnWord(BitCount,Width) + ByteAlignOnWord(1,Width))*Height)
    ' Icon sizes are limited to maximum dimensions of 256x256
    
    On Error GoTo Catch_Exception
    tDC = GetDC(0&)
    With BHI.bmiHeader
        .biSize = 40&
        If ICI.hbmColor = 0& Then  ' black and white icon (rare, but so easy)
            If GetDIBits(tDC, ICI.hbmMask, 0, 0&, ByVal 0&, BHI, 0&) Then
                .biClrUsed = 2&                 ' should be filled in, but ensure it is so
                .biClrImportant = .biClrUsed    ' should be filled in, but ensure it is so
                .biCompression = 0&
                .biSizeImage = 0&
                BHI.bmiColors(2) = vbWhite      ' set 2nd palette entry to white
                ' size array to the entire icon file format, includes Icon Directory structure, bitmap header, palette, & mask
                ReDim OutArray(0 To ByteAlignOnWord(1, .biWidth) * .biHeight + 69&)
                ' this next call gets the entire icon data; just need to fill in the directory a bit further down this routine
                If GetDIBits(tDC, ICI.hbmMask, 0, BHI.bmiHeader.biHeight, OutArray(70), BHI, 0&) = 0& Then
                    .biBitCount = 0
                Else
                    .biClrUsed = 2&                 ' fill in; last GetDIBits call erased it
                    .biClrImportant = .biClrUsed    ' fill in; last GetDIBits call erased it
                    .biHeight = .biHeight \ 2&      ' set to real height, not height*2 as is now
                End If
            End If
            DeleteObject ICI.hbmMask: ICI.hbmMask = 0& ' destroy; no longer needed
        
        Else    ' color icon vs black & white
            If GetDIBits(tDC, ICI.hbmColor, 0, 0&, ByVal 0&, BHI, 0&) Then
                .biBitCount = 32
                .biCompression = 0&
                .biClrImportant = 0&
                .biClrUsed = 0&
                .biSizeImage = 0&
                ReDim Bits(0 To .biWidth * .biHeight - 1&) ' number colors we will process
                If GetDIBits(tDC, ICI.hbmColor, 0, .biHeight, Bits(0), BHI, 0&) = 0 Then
                    .biBitCount = 0
                Else
                    ' determine if this icon can be paletted or not; fast routine for small images (256x256 or less)
                    lPrevPal = Bits(x) Xor 1&                       ' forces mismatch in loop start
                    For y = x To .biWidth * .biHeight - 1&          ' process each color
                        If Bits(y) <> lPrevPal Then
                            If (Bits(y) And &HFF000000) Then        ' uses alpha channel; 32bpp
                                .biClrImportant = 0&                ' we can abort loop; won't be paletted
                                .biBitCount = 32
                                Exit For
                                
                            ElseIf .biBitCount = 32 Then                ' continue processing else identified as potential 24bpp icon
                                palIndex = FindColor(BHI.bmiColors(), Bits(y), .biClrImportant, bNewColor) ' have we seen this color?
                                If bNewColor Then                       ' if not, add to our palette
                                    If .biClrImportant = 256& Then      ' max'd out on palette entries; treat as 24bpp
                                        .biBitCount = 24                ' but don't exit loop cause we don't know now
                                        .biClrImportant = 0&            ' if it is not a 32bpp icon
                                        
                                    Else                                ' prepare to add to our palette if new
                                        .biClrImportant = .biClrImportant + 1&
                                        If palIndex < .biClrImportant Then  ' keep our palette in ascending order for binary search
                                            CopyMemory BHI.bmiColors(palIndex + 1&), BHI.bmiColors(palIndex), (.biClrImportant - palIndex) * 4&
                                        End If
                                        BHI.bmiColors(palIndex) = Bits(y) ' add color now
                                    End If
                                End If
                            End If
                            lPrevPal = Bits(y) ' track for faster looping
                        End If
                    Next
                    maskScan = ByteAlignOnWord(1, .biWidth) ' scan width for the mask portion of this icon
                    
                    If .biClrImportant Then                ' then can be paletted
                        Select Case .biClrImportant        ' set destination bit count
                            Case Is < 3:    .biBitCount = 1
                            Case Is < 17:   .biBitCount = 4
                            Case Else:      .biBitCount = 8
                        End Select
                        pow2(0) = 1&                                ' setup a power of two lookup table
                        For y = pow2(0) To .biBitCount
                            pow2(y) = pow2(y - 1&) * 2&
                        Next
                        clrScan = ByteAlignOnWord(.biBitCount, .biWidth)    ' scan width of destination's color data
                        .biClrUsed = pow2(.biBitCount)                      ' how many palette entries we will provide
                        .biSizeImage = clrScan * .biHeight                  ' new size of color data
                        clrOffset = .biClrUsed * 4& + 62&                   ' where color data starts
                        ' size array to the entire icon file format, includes Icon Directory structure, bitmap header, palette, & mask
                        ReDim OutArray(0 To .biSizeImage + maskScan * .biHeight + clrOffset - 1&)
                        
                        lPrevPal = Bits(x) Xor 1&                   ' forces mismatch when loop starts
                        For y = x To .biHeight - 1&                 ' start packing the palette indexes into bytes
                            palShift = 8& - .biBitCount             ' 1st position of byte where palette index will be written
                            palPtr = clrOffset + y * clrScan        ' position where that byte will start for current row
                            For x = x To x + .biWidth - 1&          ' process each row of the source bitmap
                                ' locate the color in our palette & subtract one (palette is 1-based, indexes are 0-based)
                                If lPrevPal <> Bits(x) Then
                                    palIndex = FindColor(BHI.bmiColors(), Bits(x), .biClrImportant, bNewColor) - 1&
                                    lPrevPal = Bits(x) ' track for faster looping
                                End If
                                OutArray(palPtr) = OutArray(palPtr) Or (palIndex * pow2(palShift)) ' pack the index
                                If palShift = 0& Then               ' done with this byte
                                    palPtr = palPtr + 1&            ' move destination to next byte
                                    palShift = 8& - .biBitCount     ' reset the position where next index will be written
                                Else
                                    palShift = palShift - .biBitCount ' adjust position where next index will be written
                                End If
                            Next
                        Next
                    
                    Else ' 24 or 32 bit color
                            
                        .biSizeImage = ByteAlignOnWord(.biBitCount, .biWidth) * .biHeight ' size of color data
                        ' size array to the entire icon file format, includes Icon Directory structure, bitmap header
                        ReDim OutArray(0 To .biSizeImage + maskScan * .biHeight + 61&)
                        If .biBitCount = 32 Then    ' just copy the entire bitmap to our array
                            CopyMemory OutArray(62), Bits(x), .biSizeImage
                        Else
                            ' we can loop & transfer 3 of 4 bytes for each pixel or just call the API one more time
                            GetDIBits tDC, ICI.hbmColor, 0&, .biHeight, OutArray(62), BHI, 0&
                        End If
                    End If
                    Erase Bits()
                End If
            End If
        End If
    End With
        
    If BHI.bmiHeader.biBitCount Then
        With BHI.bmiHeader
            ' let's build the icon structure (22 bytes for single icon)
            ' 6 byte ICONDIRECTORY
            '   Integer: Reserved; must be zero
            '   Integer: Type. 1=Icon, 2=Cursor
            '   Integer: Count. Number ico/cur in this resource
            ' 16 BYTE ICONDIRECTORYENTRY
            ' -------- 1 of these for each ico/cur in resource. ICO entry differs from CUR entry
            '   Byte: Width; 256=0
            '   Byte: Height; 256=0
            '   Byte: Color Count; 256=0 & 16-32bit = 0
            '   Byte: Reserved; must be 0
            '   Integer: Planes; must be 1
            '   Integer: Bitcount
            '   Long: Number of bytes for this entry's ico/cur data
            '   Long: Offset into resource where ico/cur data starts
            OutArray(2) = IIf(IsIcon, 1, 2)                      ' type: icon
            OutArray(4) = 1                                      ' count
            If .biWidth < 256& Then OutArray(6) = .biWidth       ' width
            If .biHeight < 256& Then OutArray(7) = .biHeight     ' height
            If .biClrUsed < 256& Then OutArray(8) = .biClrUsed   ' color count
            OutArray(10) = IIf(IsIcon, 1, ICI.XHotSpot)           ' planes
            OutArray(12) = IIf(IsIcon, .biBitCount, ICI.YHotSpot) ' bitcount
            CopyMemory OutArray(14), CLng(UBound(OutArray) - 21&), 4& ' bytes in resource
            OutArray(18) = 22                                    ' offset into directory where BHI starts
            .biHeight = .biHeight + .biHeight                    ' icon's store height*2 in bitmap header
        End With
        ' copy the bitmap header & palette, if used
        CopyMemory OutArray(OutArray(18)), BHI, BHI.bmiHeader.biClrUsed * 4& + BHI.bmiHeader.biSize
        
        ' done with the icon directory, now to the mask portion
        If ICI.hbmMask Then
            BHI.bmiColors(1) = vbBlack: BHI.bmiColors(2) = vbWhite      ' set up black/white palette
            With BHI.bmiHeader                                          ' set up bitmapinfo header
                .biBitCount = 1
                .biClrUsed = 2&
                .biClrImportant = .biClrUsed
                .biHeight = .biHeight \ 2&
                .biSizeImage = 0&
                palPtr = UBound(OutArray) - maskScan * .biHeight + 1&    ' location where mask will be written
            End With
            GetDIBits tDC, ICI.hbmMask, 0&, BHI.bmiHeader.biHeight, OutArray(palPtr), BHI, 0& ' get the mask
        End If
        SaveHICONtoArray = True
    End If
    
Catch_Exception:
    ReleaseDC 0&, tDC
    If ICI.hbmColor Then DeleteObject ICI.hbmColor
    If ICI.hbmMask Then DeleteObject ICI.hbmMask
    
End Function

Private Function FindColor(ByRef PaletteItems() As Long, ByVal Color As Long, ByVal Count As Long, ByRef isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    
    ' [in] PaletteItems(). Long Array to search within. Array must be 1-bound
    ' [in] Color. A value to search for. Order is always ascending
    ' [in] Count. Number of items in PaletteItems() to compare against
    ' [out] isNew. If Color not found, isNew is True else False
    ' [out] Return value: The Index where Color was found or where the new Color should be inserted

    Dim ub As Long, lb As Long
    Dim newIndex As Long
    
    If Count = 0& Then
        FindColor = 1&
        isNew = True
        Exit Function
    End If
    
    ub = Count
    lb = 1&
    
    Do Until lb > ub
        newIndex = lb + ((ub - lb) \ 2&)
        Select Case PaletteItems(newIndex) - Color
        Case 0& ' match found
            Exit Do
        Case Is > 0& ' new color is lower in sort order
            ub = newIndex - 1&
        Case Else ' new color is higher in sort order
            lb = newIndex + 1&
        End Select
    Loop

    If lb > ub Then  ' color was not found
            
        If Color > PaletteItems(newIndex) Then newIndex = newIndex + 1&
        isNew = True
        
    Else
        isNew = False
    End If
    
    FindColor = newIndex

End Function

Private Function ByteAlignOnWord(ByVal bitDepth As Long, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    ByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

