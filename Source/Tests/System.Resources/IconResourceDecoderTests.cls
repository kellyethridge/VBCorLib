VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconResourceDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' IconResourceDecoderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Decode_WithNonIconResourceType_ThrowsArgumentException()
    Dim Decoder As New IconResourceDecoder
    Dim Key As ResourceKey
    Set Key = NewResourceKey(101, ResourceTypes.StringResource, 1033)
    
    On Error Resume Next
    Decoder.Decode Key, NewBytes()
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Decode_WithIconData_ReturnsValidStdPicture()
    Dim Decoder As New IconResourceDecoder
    Dim Data() As Byte
    Dim Key As ResourceKey
    Dim ActualCount As Long
    Dim Resource As Variant
    
    Set Key = NewResourceKey(101, ResourceTypes.IconResource, 1033)
    Data = GetIconData
    ActualCount = Decoder.Decode(Key, Data)
    Decoder.GetResource Key, Resource
    
    Assert.That ActualCount, Iz.EqualTo(1)
    Assert.That TypeOf Resource Is StdPicture, Iz.True
    Dim Pic As StdPicture
    Set Pic = Resource
    Assert.That Pic.Type, Iz.EqualTo(vbPicTypeIcon)
    Assert.That Pic.Handle, Iz.Not.EqualTo(0)
End Sub

Public Sub Decode_WithInvalidIconData_ReturnsNothing()
    Dim Decoder As New IconResourceDecoder
    Dim Data() As Byte
    Dim Key As ResourceKey
    Dim ActualCount As Long
    Dim Resource As Variant
    
    Set Key = NewResourceKey(101, ResourceTypes.IconResource, 1033)
    Data = NewBytes(1, 2, 3, 4)
    ActualCount = Decoder.Decode(Key, Data)
    Decoder.GetResource Key, Resource
    
    Assert.That ActualCount, Iz.EqualTo(0)
End Sub


Private Function GetIconData() As Byte()
    Dim Encoder As New IconResourceEncoder
    Dim Key As ResourceKey
    Dim Data() As Byte
    Encoder.Encode LoadIcon, "Test"
    Encoder.GetEncodedResource Key, Data
    GetIconData = Data
End Function
