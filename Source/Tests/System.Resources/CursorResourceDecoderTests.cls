VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CursorResourceDecoderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Decode_WithNonCursorResourceType_ThrowsArgumentException()
    Dim Decoder As New CursorResourceDecoder
    Dim Key As ResourceKey
    Set Key = NewResourceKey(101, ResourceTypes.StringResource, 1033)
    
    On Error Resume Next
    Decoder.Decode Key, NewBytes()
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Decode_WithCursorData_ReturnsValidStdPicture()
    Dim Decoder As New CursorResourceDecoder
    Dim Data() As Byte
    Dim Key As ResourceKey
    Dim ActualCount As Long
    Dim Resource As Variant
    
    Set Key = NewResourceKey(101, ResourceTypes.CursorResource, 1033)
    Data = GetCursorData
    ActualCount = Decoder.Decode(Key, Data)
    Decoder.GetResource Key, Resource
    
    Assert.That ActualCount, Iz.EqualTo(1)
    Assert.That TypeOf Resource Is StdPicture, Iz.True
    Dim Pic As StdPicture
    Set Pic = Resource
    Assert.That Pic.Type, Iz.EqualTo(vbPicTypeIcon)
    Assert.That Pic.Handle, Iz.Not.EqualTo(0)
End Sub

Public Sub Decode_WithInvalidCursorData_ReturnsNothing()
    Dim Decoder As New CursorResourceDecoder
    Dim Data() As Byte
    Dim Key As ResourceKey
    Dim ActualCount As Long
    Dim Resource As Variant
    
    Set Key = NewResourceKey(101, ResourceTypes.CursorResource, 1033)
    Data = NewBytes(1, 2, 3, 4)
    ActualCount = Decoder.Decode(Key, Data)
    Decoder.GetResource Key, Resource
    
    Assert.That ActualCount, Iz.EqualTo(0)
End Sub

Private Function GetCursorData() As Byte()
    Dim Encoder As New CursorResourceEncoder
    Dim Key As ResourceKey
    Dim Data() As Byte
    Encoder.Encode LoadCursor, "Test"
    Encoder.GetEncodedResource Key, Data
    GetCursorData = Data
End Function
