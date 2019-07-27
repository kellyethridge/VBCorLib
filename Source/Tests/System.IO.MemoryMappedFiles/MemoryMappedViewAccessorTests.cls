VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedViewAccessorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' MemoryMappedViewAccessorTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private mIOHelper As IOHelper

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO.MemoryMappedFiles"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ReadByte_WithInvalidPosition_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 100
            
        Case "ReadByte_WhenCalled_ReturnsExpected"
            Test.Use 0, 0
            Test.Use 1, 1
            Test.Use 99, 99
            
        Case "ReadByte_WithOneByteOffseet_ReturnsExpected"
            Test.Use 0, 1
            Test.Use 1, 2
            Test.Use 98, 99
        
        Case "ReadBoolean_WhenCalled_ReturnsExpected"
            Test.Use 0, False
            Test.Use 1, True
            Test.Use 2, True
            
        Case "Read_WithInvalidStructure_ThrowsArgumentException"
            Dim s As String
            s = "hello"
            Test.Use s
            Test.Use NewLongs()
            Test.Use Array()
            Test.Use Nothing
            
            Dim Obj As New Collection
            Test.Use Obj
            
            Dim SS As StringStructure
            Test.Use SS
            
            Dim CS As InvalidComplexStructure
            Test.Use CS
            
        Case "Read_WhenCalledWithDataType_FillsDataType"
            Test.Use CByte(0), 1
            Test.Use CInt(0), 513
            Test.Use CLng(0), 67305985
            Test.Use False, True
            Test.Use CCur(0), 57843769575230.7201@
            Test.Use CDate(0), #12/31/1833 1:46:15 AM#
            
        Case "ReadArray_WhenCalled_FillsArray"
            Test.Use CorArray.CreateInstance(vbByte, 2), Cor.NewBytes(0, 1)
            Test.Use CorArray.CreateInstance(vbBoolean, 2), Cor.NewBooleans(False, True)
            Test.Use CorArray.CreateInstance(vbLong, 3), Cor.NewLongs(50462976, 117835012, 185207048)
            Test.Use CorArray.CreateInstance(vbDate, 3), Cor.NewDates(#10/5/1604 12:31:31 AM#, #8/28/3438 10:29:21 AM#, #7/19/5272 8:27:12 PM#)
            
        Case "WriteValue_WhenCalled_WritesExpected"
            Test.Use CByte(255), Cor.NewBytes(255)
            Test.Use CInt(32767), Cor.NewBytes(&HFF, &H7F)
            Test.Use CLng(&H7FFFFFFF), Cor.NewBytes(&HFF, &HFF, &HFF, &H7F)
            Test.Use 922337203685477.5807@, Cor.NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &H7F)
            Test.Use #12/31/1833 1:46:15 AM#, Cor.NewBytes(&H80, &HCD, &HDF, &H3, &H5, &H6, &H7, &H8)
            Test.Use True, Cor.NewBytes(1)
            Test.Use CDec(123456.789012), Cor.NewBytes(&H14, &H1A, &H99, &HBE, &H1C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0)
            Test.Use CSng(12.345), Cor.NewBytes(&H1F, &H85, &H45, &H41)
            Test.Use CDbl(12.345), Cor.NewBytes(&H71, &H3D, &HA, &HD7, &HA3, &HB0, &H28, &H40)
            Test.Use Cor.NewInt64(&H90123456, &H12345678), Cor.NewBytes(&H56, &H34, &H12, &H90, &H78, &H56, &H34, &H12)
            
        Case "WriteArray_WithArray_WritesExpected"
            Test.Use NewBytes(3, 2, 1), NewBytes(3, 2, 1)
            Test.Use NewIntegers(1, 2, 3), NewBytes(1, 0, 2, 0, 3, 0)
            Test.Use NewLongs(1, 2, 3), NewBytes(1, 0, 0, 0, 2, 0, 0, 0, 3, 0, 0, 0)
            Test.Use NewCurrencies(1, 2, 3), NewBytes(&H10, &H27, 0, 0, 0, 0, 0, 0, &H20, &H4E, 0, 0, 0, 0, 0, 0, &H30, &H75, 0, 0, 0, 0, 0, 0)
            Test.Use NewDates(#12/31/1833 1:46:15 AM#, #12/31/1833 1:46:15 AM#), Cor.NewBytes(&H80, &HCD, &HDF, &H3, &H5, &H6, &H7, &H8, &H80, &HCD, &HDF, &H3, &H5, &H6, &H7, &H8)
            Test.Use NewBooleans(True, False, True), NewBytes(1, 0, 1)
            Test.Use Array(CDec(123456.789012), CLng(&H55555555)), Cor.NewBytes(&H14, &H1A, &H99, &HBE, &H1C, 0, 0, 0, 0, 0, 0, 0, 0, 0, 6, 0, &H55, &H55, &H55, &H55)
            Test.Use NewSingles(12.345, 12.345), NewBytes(&H1F, &H85, &H45, &H41, &H1F, &H85, &H45, &H41)
            Test.Use NewDoubles(12.345, 12.345), Cor.NewBytes(&H71, &H3D, &HA, &HD7, &HA3, &HB0, &H28, &H40, &H71, &H3D, &HA, &HD7, &HA3, &HB0, &H28, &H40)
            
            Dim Int64s(0 To 1) As Int64
            Int64s(0) = Cor.NewInt64(&H90123456, &H12345678)
            Int64s(1) = Cor.NewInt64(&H90123456, &H12345678)
            Test.Use Int64s, Cor.NewBytes(&H56, &H34, &H12, &H90, &H78, &H56, &H34, &H12, &H56, &H34, &H12, &H90, &H78, &H56, &H34, &H12)
        
    End Select
End Sub

Public Sub Setup()
    Set mIOHelper = New IOHelper
End Sub

Public Sub Teardown()
    Set mIOHelper = Nothing
End Sub

Public Sub ReadByte_WithInvalidPosition_ThrowsArgumentOutOfRangeException(ByVal InvalidPosition As Long)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    
    Set MappedFile = MakeEmptyMappedFile
    Set Accessor = MappedFile.CreateViewAccessor(Size:=100)

    On Error Resume Next
    Accessor.ReadByte InvalidPosition

    AssertArgumentOutOfRangeException Err, "Position"
End Sub

Public Sub ReadByte_WhenCalled_ReturnsExpected(ByVal Position As Long, ByVal Expected As Byte)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Byte
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadByte(Position)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadByte_WithOneByteOffseet_ReturnsExpected(ByVal Position As Long, ByVal Expected As Byte)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Byte
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor(1)
    
    Actual = Accessor.ReadByte(Position)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadBoolean_WhenCalled_ReturnsExpected(ByVal Position As Long, ByVal Expected As Boolean)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Boolean
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    Actual = Accessor.ReadBoolean(Position)

    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadChar_WhenCalled_Returns256()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Integer
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    Actual = Accessor.ReadChar(0)

    Assert.That Actual, Iz.EqualTo(256)
End Sub

Public Sub ReadInt16_WhenReadingFromBeginning_Returns256()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Integer
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    Actual = Accessor.ReadInt16(0)

    Assert.That Actual, Iz.EqualTo(256)
End Sub

Public Sub ReadInt32_WhenReadingFromBeginning_Returns50462976()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Long
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadInt32(0)

    Assert.That Actual, Iz.EqualTo(50462976)
End Sub

Public Sub ReadCurrency_WhenReadingFromBeginning_Returns50609752291423Point0528()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Currency
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadCurrency(0)

    Assert.That Actual, Iz.EqualTo(50609752291423.0528@)
End Sub

Public Sub ReadDouble_WhenReadingFromBeginning_ReturnsExpected()
    Const Expected As Double = 7.94992889512736E-275
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Double
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    Actual = Accessor.ReadDouble(0)

    Assert.That Actual, Iz.EqualTo(Expected).Within(1E-289)
End Sub

Public Sub ReadSingle_WhenReadingFromBeginning_ReturnsExpected()
    Const Expected As Single = 3.820471E-37
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Single
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadSingle(0)

    Assert.That Actual, Iz.EqualTo(Expected).Within(1E-43)
End Sub

Public Sub ReadDecimal_WhenReadingFromPosition88_ReturnsExpected()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Variant
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadDecimal(88)

    Assert.That Actual, Iz.EqualTo(CDec("30757950549943697627242191192"))
End Sub

Public Sub ReadInt64_WhenReadingFromBeginning_ReturnsExpected()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Int64
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadInt64(0)

    Assert.That Actual, Iz.EqualTo(NewInt64(&H3020100, &H7060504))
End Sub

Public Sub ReadInt64_WhenUsingSuppliedFileStream_ReadsExpected()
    Dim FileStream  As FileStream
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Int64
    
    mIOHelper.CreateBinaryFile Data:=MakeData
    Set FileStream = Cor.NewFileStream(mIOHelper.TempFileName, FileMode.OpenExisting)
    Set MappedFile = MemoryMappedFile.CreateFromFileStream(FileStream, vbNullString, 0, MemoryMappedFileAccess.ReadWrite, True)
    Set Accessor = MappedFile.CreateViewAccessor
    
    Actual = Accessor.ReadInt64(0)

    Assert.That Actual, Iz.EqualTo(NewInt64(&H3020100, &H7060504))
End Sub

Public Sub Read_WhenCalledWithStructure_FillsStructure()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As Int64
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Accessor.Read 0, Actual

    Assert.That Actual, Iz.EqualTo(NewInt64(&H3020100, &H7060504))
End Sub

Public Sub Read_WhenCalledWithComplexStructure_FillsStructure()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual      As ComplexStructure
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    Accessor.Read 0, Actual

    Dim Expected As ComplexStructure
    Expected.Value = &H3020100
    Expected.FixedLong(0) = &H7060504
    Expected.Inner.Value = &HB0A0908
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Read_WhenCalledWithDataType_FillsDataType(ByRef DataType As Variant, ByRef Expected As Variant)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    Accessor.Read 1, DataType

    Assert.That DataType, Iz.EqualTo(Expected)
End Sub

Public Sub Read_WithInvalidStructure_ThrowsArgumentException(ByRef InvalidStructure As Variant)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    On Error Resume Next
    Accessor.Read 0, InvalidStructure

    AssertArgumentException Err, "Structure"
End Sub

Public Sub ReadArray_WhenCalledWithNonArray_ThrowsArgumentException()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Invalid     As Long
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    On Error Resume Next
    Accessor.ReadArray 0, Invalid, 0, 1

    AssertArgumentException Err
End Sub

Public Sub ReadArray_WhenCalledWithNonUTDArray_ThrowsArgumentException()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Invalid(0)  As Object
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    On Error Resume Next
    Accessor.ReadArray 0, Invalid, 0, 1

    AssertArgumentException Err
End Sub

Public Sub ReadArray_WithUTDArray_FillsArray()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual(1)   As Int64
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    Accessor.ReadArray 0, Actual, 0, 2

    Assert.That Actual, Iz.EqualTo(Array(NewInt64(&H3020100, &H7060504), NewInt64(&HB0A0908, &HF0E0D0C)))
End Sub

Public Sub ReadArray_WhenCalled_FillsArray(ByRef Actual As Variant, ByRef Expected As Variant)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim ActualCount As Long
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor
    
    ActualCount = Accessor.ReadArray(0, Actual, 0, CorArray.Length(Actual))

    Assert.That ActualCount, Iz.EqualTo(CorArray.Length(Expected)), "#1"
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadArray_WhenRequestingMoreThanAvailable_FillsOnlyAvailable()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual(0 To 999) As Int64
    Dim ActualCount As Long
    
    Set MappedFile = MakeMappedFileWithData
    Set Accessor = MappedFile.CreateViewAccessor

    ActualCount = Accessor.ReadArray(0, Actual, 0, 1000)

    Assert.That ActualCount, Iz.EqualTo(512), "#1"
    Assert.That Actual(12), Iz.EqualTo(Cor.NewInt64(&H63626160, 0)), "#2"
    Assert.That Actual(13), Iz.EqualTo(Cor.NewInt64(0, 0)), "#3" ' uses remaining allocated page which is zero-filled.
End Sub

Public Sub WriteValue_WhenCalled_WritesExpected(ByRef Value As Variant, ByRef Expected() As Byte)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual()    As Byte
    
    Set MappedFile = MakeEmptyMappedFile
    Set Accessor = MappedFile.CreateViewAccessor
    
    Accessor.WriteValue 0, Value
    
    Set MappedFile = Nothing
    Set Accessor = Nothing
    Actual = ReadBytes(0, CorArray.Length(Expected))
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteArray_WhenCannotWrite_ThrowsNotSupportedException()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    
    Set MappedFile = MakeEmptyMappedFile
    Set Accessor = MappedFile.CreateViewAccessor(Access:=MemoryMappedFileAccess.ReadOnly)

    On Error Resume Next
    Accessor.WriteArray 0, Cor.NewBytes(0), 0, 1

    AssertNotSupportedException Err
End Sub

Public Sub WriteArray_WithNegativePosition_ThrowsArgumentOutOfRangeException()
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    
    Set MappedFile = MakeEmptyMappedFile
    Set Accessor = MappedFile.CreateViewAccessor

    On Error Resume Next
    Accessor.WriteArray -1, Cor.NewBytes(0), 0, 1

    AssertArgumentOutOfRangeException Err, "Position"
End Sub

Public Sub WriteArray_WithArray_WritesExpected(ByRef Value As Variant, ByRef Expected As Variant)
    Dim MappedFile  As MemoryMappedFile
    Dim Accessor    As MemoryMappedViewAccessor
    Dim Actual()    As Byte
    
    Set MappedFile = MakeEmptyMappedFile
    Set Accessor = MappedFile.CreateViewAccessor

    Accessor.WriteArray 0, Value, 0, CorArray.Length(Value)

    Set MappedFile = Nothing
    Set Accessor = Nothing
    Actual = ReadBytes(0, CorArray.Length(Expected))
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Private Function MakeEmptyMappedFile(Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    Dim Data(0 To 99) As Byte
    Set MakeEmptyMappedFile = MakeMappedFile(Data)
End Function

Private Function MakeMappedFileWithData(Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    Set MakeMappedFileWithData = MakeMappedFile(MakeData)
End Function

Private Function MakeMappedFile(ByRef Data() As Byte, Optional ByVal Access As MemoryMappedFileAccess = MemoryMappedFileAccess.ReadWrite) As MemoryMappedFile
    mIOHelper.CreateBinaryFile "File.txt", Data
    Set MakeMappedFile = MemoryMappedFile.CreateFromFile(mIOHelper.TempFileName, Access:=Access)
End Function

Private Function ReadBytes(ByVal Offset As Long, ByVal Count As Long) As Byte()
    With Cor.NewFileStream(mIOHelper.TempFileName, OpenExisting)
        Dim Data() As Byte
        ReDim Data(0 To Count - 1)
        .Position = Offset
        .ReadBlock Data, 0, Count
        ReadBytes = Data
    End With
End Function

Private Function MakeData() As Byte()
    Dim Data(0 To 99) As Byte
    Dim i As Long
    
    For i = LBound(Data) To UBound(Data)
        Data(i) = i Mod 256
    Next

    MakeData = Data
End Function

