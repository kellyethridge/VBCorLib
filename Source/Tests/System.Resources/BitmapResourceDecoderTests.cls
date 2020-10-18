VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitmapResourceDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BitmapResourceDecoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Decode_WithBitmapData_ReturnsValidStdPicture"
            Test.Use False
            Test.Use True
    End Select
End Sub

Public Sub Decode_WithNonBitmapResourceType_ThrowsArgumentException()
    Dim Decoder As New BitMapResourceDecoder
    Dim Key As ResourceKey
    Set Key = NewResourceKey(101, ResourceTypes.StringResource, 1033)
    
    On Error Resume Next
    Decoder.Decode Key, NewBytes()
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Decode_WithBitmapData_ReturnsValidStdPicture(ByVal ExcludeFileHeader As Boolean)
    Dim Decoder As New BitMapResourceDecoder
    Dim Data() As Byte
    Dim Key As ResourceKey
    Dim ActualCount As Long
    Dim Resource As Variant
    
    Set Key = NewResourceKey(101, ResourceTypes.BitmapResource, 1033)
    Data = GetBitmapData(ExcludeFileHeader)
    ActualCount = Decoder.Decode(Key, Data)
    Decoder.GetResource Key, Resource
    
    Assert.That ActualCount, Iz.EqualTo(1)
    Assert.That TypeOf Resource Is StdPicture, Iz.True
    Dim Pic As StdPicture
    Set Pic = Resource
    Assert.That Pic.Type, Iz.EqualTo(vbPicTypeBitmap)
    Assert.That Pic.Handle, Iz.Not.EqualTo(0)
End Sub

Private Function GetBitmapData(Optional ByVal ExcludeFileHeader As Boolean) As Byte()
    Dim Stream As Stream
    Set Stream = File.OpenRead(MakeResourcePath("balloon.bmp"))
    
    If ExcludeFileHeader Then
        Stream.Position = 14
    End If
    
    GetBitmapData = Cor.NewBinaryReader(Stream).ReadBytes(Stream.Length - Stream.Position)
End Function


