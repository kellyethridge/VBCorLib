VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceGroupEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CursorResourceGroupEncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Encode_WithUnsupportedValues_ReturnsFalse"
            Test.Use Nothing
            Test.Use New CursorResourceEncoder
            Test.Use Cor.NewPictureResourceGroup(IconGroup)
            
    End Select
End Sub

Public Sub Encode_WithUnsupportedValues_ReturnsFalse(ByRef Value As Variant)
    Dim Encoder As New CursorResourceGroupEncoder
    
    Assert.That Encoder.Encode(Value, 1), Iz.False
End Sub

Public Sub Encode_WithEmptyPictureResourceGroup_ReturnsFalse()
    Dim Encoder As New CursorResourceGroupEncoder
    Dim Group   As PictureResourceGroup
    Set Group = Cor.NewPictureResourceGroup(CursorGroup)
    
    Assert.That Encoder.Encode(Group, 1), Iz.False
End Sub

Public Sub Encode_WithPictureResourceGroup_ReturnsTrue()
    Dim Encoder As New CursorResourceGroupEncoder
    Dim Group   As PictureResourceGroup
    Set Group = Cor.NewPictureResourceGroup(CursorGroup)
    Group.Add MakePictureResourceInfo
    
    Assert.That Encoder.Encode(Group, 1), Iz.True
End Sub

Public Sub Encode_WhenCalled_EncodesResource()
    Dim Encoder As New CursorResourceGroupEncoder
    Dim Group   As PictureResourceGroup
    Set Group = Cor.NewPictureResourceGroup(CursorGroup)
    Group.Add MakePictureResourceInfo
    Encoder.Encode Group, 1
    
    Dim ActualKey As ResourceKey
    Dim ActualData() As Byte
    Assert.That Encoder.GetEncodedResource(ActualKey, ActualData), Iz.True
    
    Assert.That ActualKey.ResourceName, Iz.EqualTo(1)
    Dim ExpectedData() As Byte
    ExpectedData = Cor.NewBytes(0, 0, 2, 0, 1, 0, 1, 0, 2, 0, 3, 0, 4, 0, 5, 0, 0, 0, 6, 0)
    Assert.That ActualData, Iz.EqualTo(ExpectedData)
End Sub

Private Function MakePictureResourceInfo() As PictureResourceInfo
    Dim Info As New FakePictureResourceInfo
    Info.ResourceType = CursorResource
    Info.Width = 1
    Info.Height = 2
    Info.Planes = 3
    Info.BitCount = 4
    Info.Size = 5
    Info.ResourceId = 6
    Set MakePictureResourceInfo = Info
End Function
