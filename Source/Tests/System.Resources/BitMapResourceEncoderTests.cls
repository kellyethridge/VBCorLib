VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitMapResourceEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BitMapResourceEncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Encode_WithUnsupportedValue_ReturnsFalse"
            Test.Use "abc"
            Test.Use Nothing
            Test.Use New Collection
            Test.Use LoadPicture("checkmrk.ico")
            Test.Use LoadPicture("normal01.cur")
            
    End Select
End Sub

Public Sub Encode_WithUnsupportedValue_ReturnsFalse(ByRef UnsupportedValue As Variant)
    Dim Encoder As New BitMapResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(UnsupportedValue, "Test")
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Encode_WithBitMap_Encodes()
    Dim Encoder As New BitMapResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(LoadPicture("balloon.bmp"), "Test")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetEncodedResource_WhenCalled_ReturnsKeyAndValue()
    Dim Encoder As New BitMapResourceEncoder
    Encoder.Encode LoadPicture("balloon.bmp"), "Test"

    Dim Key As ResourceKey
    Dim Value As Variant
    Assert.That Encoder.GetEncodedResource(Key, Value), Iz.True
    
    Assert.That Key.ResourceName, Iz.EqualTo("Test")
    Assert.That Key.ResourceType, Iz.EqualTo(ResourceTypes.BitmapResource)
    Dim DataSize As Long
    DataSize = FileLen(MakeFilePath("balloon.bmp")) - 14 ' exclude the bitmap file header.
    Assert.That CorArray.Length(Value), Iz.EqualTo(DataSize)
    Assert.That Value(0), Iz.EqualTo(40)    ' just check for correct bitmap data header size.
End Sub

Private Function LoadPicture(ByVal FileName As String) As StdPicture
    Set LoadPicture = VB.LoadPicture(MakeFilePath(FileName))
End Function

Private Function MakeFilePath(ByVal FileName As String) As String
    MakeFilePath = Path.Combine(App.Path, CorString.Format("System.Resources/{0}", FileName))
End Function
