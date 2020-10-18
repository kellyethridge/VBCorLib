VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CursorResourceEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CursorResourceEncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_CURSOR As Long = 2
Private Const LR_LOADFROMFILE As Long = &H10

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
            Case "Encode_WithUnsupportedValue_ReturnsFalse"
            Test.Use "abc"
            Test.Use Nothing
            Test.Use New Collection
            Test.Use LoadIcon
            Test.Use LoadBitMap
    
    End Select
End Sub

Public Sub Encode_WithUnsupportedValue_ReturnsFalse(ByRef UnsupportedValue As Variant)
    Dim Encoder As New CursorResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(UnsupportedValue, "Test")
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Encode_WithStdPictureCursor_ReturnsTrue()
    Dim Encoder As New CursorResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(LoadCursor, "Test")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetEncodedResource_WhenEncodedFromStdPicture_ReturnsKeyAndValue()
    Dim Encoder As New CursorResourceEncoder
    Encoder.Encode LoadCursor, "Test"

    Dim Key As ResourceKey
    Dim Value() As Byte
    Assert.That Encoder.GetEncodedResource(Key, Value), Iz.True
    
    Assert.That Key.ResourceName, Iz.EqualTo("Test")
    Assert.That Key.ResourceType, Iz.EqualTo(ResourceTypes.CursorResource)
    Dim DataSize As Long
    DataSize = FileLen(MakeResourcePath("normal01.cur")) - 22 + 4 ' exclude the bitmap file header but add 4 bytes for X and Y hot-spot offsets.
    Assert.That CorArray.Length(Value), Iz.EqualTo(DataSize)
    Assert.That Value(4), Iz.EqualTo(40)    ' just check for correct bitmap data header size.
End Sub

Public Sub Encode_WithPictureHandle_ReturnsTrue()
    Dim Encoder As New CursorResourceEncoder
    Dim Handle As Long
    
    Handle = LoadImage(0, MakeResourcePath("normal01.cur"), IMAGE_CURSOR, 0, 0, LR_LOADFROMFILE)
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(Handle, "Test")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetEncodedResource_WhenEncodedFromPictureHandle_ReturnsKeyAndValue()
    Dim Encoder As New CursorResourceEncoder
    Dim Handle As Long
    
    Handle = LoadImage(0, MakeResourcePath("normal01.cur"), IMAGE_CURSOR, 0, 0, LR_LOADFROMFILE)
    Encoder.Encode Handle, "Test"

    Dim Key As ResourceKey
    Dim Value() As Byte
    Assert.That Encoder.GetEncodedResource(Key, Value), Iz.True
    
    Assert.That Key.ResourceName, Iz.EqualTo("Test")
    Assert.That Key.ResourceType, Iz.EqualTo(ResourceTypes.CursorResource)
    Dim DataSize As Long
    DataSize = FileLen(MakeResourcePath("normal01.cur")) - 22 + 4 ' exclude the bitmap file header but add 4 bytes for X and Y hot-spot offsets.
    Assert.That CorArray.Length(Value), Iz.EqualTo(DataSize)
    Assert.That Value(4), Iz.EqualTo(40)    ' just check for correct bitmap data header size.
End Sub

