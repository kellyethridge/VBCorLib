VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureResourceInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' PictureResourceInfoTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub NewPictureResourceGroupInfo_WithNothingResource_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewPictureResourceInfo 1, Nothing
    
    AssertArgumentNullException Err, "Resource"
End Sub

Public Sub NewPictureResourceGroupInfo_WithBitmapResource_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewPictureResourceInfo 1, LoadBitMap
    
    AssertArgumentException Err, "Resource"
End Sub

Public Sub NewPictureResourceGroupInfo_WithIcon_SetsProperties()
    Dim Info As PictureResourceInfo

    Set Info = Cor.NewPictureResourceInfo(1, LoadIcon)

    Assert.That Info.ResourceId, Iz.EqualTo(1)
    Assert.That Info.ResourceType, Iz.EqualTo(ResourceTypes.IconResource)
    Assert.That Info.Height, Iz.GreaterThan(0)
    Assert.That Info.Width, Iz.GreaterThan(0)
    Assert.That Info.Colors, Iz.GreaterThan(0)
    Assert.That Info.BitCount, Iz.GreaterThan(0)
    Assert.That Info.Planes, Iz.GreaterThan(0)
    Assert.That Info.Size, Iz.GreaterThan(0)
End Sub

Public Sub NewPictureResourceGroupInfo_WithCursor_SetsProperties()
    Dim Info As PictureResourceInfo

    Set Info = Cor.NewPictureResourceInfo(1, LoadCursor)

    Assert.That Info.ResourceId, Iz.EqualTo(1)
    Assert.That Info.ResourceType, Iz.EqualTo(ResourceTypes.CursorResource)
    Assert.That Info.Height, Iz.GreaterThan(0)
    Assert.That Info.Width, Iz.GreaterThan(0)
    Assert.That Info.Colors, Iz.GreaterThan(0)
    Assert.That Info.BitCount, Iz.EqualTo(0)
    Assert.That Info.Planes, Iz.GreaterThan(0)
    Assert.That Info.Size, Iz.GreaterThan(0)
End Sub

