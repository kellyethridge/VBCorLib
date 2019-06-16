VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringResourceDecoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StringResourceDecoderTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Public Sub Decode_WithNonStringResourceType_ThrowsArgumentException()
    Dim Decoder As New StringResourceDecoder
    Dim Key As ResourceKey
    Set Key = NewResourceKey("Test", ResourceTypes.IconResource, 1033)
    
    On Error Resume Next
    Decoder.Decode Key, NewBytes()
    
    AssertArgumentException Err, "Key"
End Sub

Public Sub Decode_WithStringData_DecodesData()
    Dim Decoder As New StringResourceDecoder
    Dim Key As ResourceKey
    Dim Count As Long
    Dim Key1 As ResourceKey
    Dim Key2 As ResourceKey
    Dim String1 As String
    Dim String2 As String
    
    Set Key = NewResourceKey(1, ResourceTypes.StringResource, 1033)
    Count = Decoder.Decode(Key, NewBytes(0, 0, 1, 0, &H61, 0, 1, 0, &H62, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
    Decoder.GetResource Key1, String1
    Decoder.GetResource Key2, String2
    
    Assert.That Count, Iz.EqualTo(2), "Count"
    Assert.That Key1.ResourceName, Iz.EqualTo(1), "Key1"
    Assert.That String1, Iz.EqualTo("a"), "String1"
    Assert.That Key2.ResourceName, Iz.EqualTo(2), "Key2"
    Assert.That String2, Iz.EqualTo("b"), "String2"
End Sub
