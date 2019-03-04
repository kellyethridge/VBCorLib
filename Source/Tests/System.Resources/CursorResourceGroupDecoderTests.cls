VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceGroupDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CursorResourceGroupDecoderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Decode_WithData_ReturnsOne()
    Dim Decoder As New CursorResourceGroupDecoder
    Dim Key As ResourceKey
    Dim Data() As Byte
    
    Data = Cor.NewBytes(0, 0, 2, 0, 1, 0, 1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 0, 0, 6, 0)
    Set Key = Cor.NewResourceKey(101, ResourceTypes.GroupCursor, 1033)
    
    Dim Actual As Long
    Actual = Decoder.Decode(Key, Data)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub GetResource_WhenCalled_ReturnsPictureResourceGroup()
    Dim Decoder As New CursorResourceGroupDecoder
    Dim Key As ResourceKey
    Dim Data() As Byte
    
    Data = Cor.NewBytes(0, 0, 2, 0, 1, 0, 1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 0, 0, 6, 0)
    Set Key = Cor.NewResourceKey(101, ResourceTypes.GroupCursor, 1033)
    Decoder.Decode Key, Data

    Dim Resource As PictureResourceGroup
    Assert.That Decoder.GetResource(Key, Resource), Iz.True
    
    Assert.That TypeOf Resource Is PictureResourceGroup, Iz.True
    Assert.That Resource.Count, Iz.EqualTo(1)
    Assert.That Resource.GroupType, Iz.EqualTo(PictureGroupTypes.CursorGroup)
    Dim Info As PictureResourceInfo
    Set Info = Resource(0)
    Assert.That Info.ResourceType, Iz.EqualTo(CursorResource)
    Assert.That Info.Width, Iz.EqualTo(1)
    Assert.That Info.Height, Iz.EqualTo(2)
    Assert.That Info.Planes, Iz.EqualTo(3)
    Assert.That Info.BitCount, Iz.EqualTo(4)
    Assert.That Info.Size, Iz.EqualTo(5)
    Assert.That Info.ResourceId, Iz.EqualTo(6)
    Assert.That Info.Colors, Iz.EqualTo(16)
    Assert.That Decoder.GetResource(Key, Resource), Iz.False
End Sub
