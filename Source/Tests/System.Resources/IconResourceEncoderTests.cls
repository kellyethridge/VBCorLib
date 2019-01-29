VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconResourceEncoderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' IconResourceEncoderTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_LOADFROMFILE As Long = &H10
Private Const IMAGE_ICON As Long = 1
Private Const IconFileName As String = "checkmrk.ico"

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Resources"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
            Case "Encode_WithUnsupportedValue_ReturnsFalse"
            Test.Use "abc"
            Test.Use Nothing
            Test.Use New Collection
            Test.Use LoadCursor
            Test.Use LoadBitMap
    
    End Select
End Sub

Public Sub Encode_WithUnsupportedValue_ReturnsFalse(ByRef UnsupportedValue As Variant)
    Dim Encoder As New IconResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(UnsupportedValue, "Test")
    
    Assert.That Actual, Iz.False
End Sub

Public Sub Encode_WithStdPictureIcon_ReturnsTrue()
    Dim Encoder As New IconResourceEncoder
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(LoadIcon, "Test")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetEncodedResource_WhenEncodedFromStdPicture_ReturnsKeyAndValue()
    Dim Encoder As New IconResourceEncoder
    Encoder.Encode LoadIcon, "Test"

    Dim Key As ResourceKey
    Dim Value() As Byte
    Assert.That Encoder.GetEncodedResource(Key, Value), Iz.True
    
    Assert.That Key.ResourceName, Iz.EqualTo("Test")
    Assert.That Key.ResourceType, Iz.EqualTo(ResourceTypes.IconResource)
    Assert.That Value(0), Iz.EqualTo(40)    ' just check for correct bitmap data header size.
End Sub

Public Sub Encode_WithPictureHandle_ReturnsTrue()
    Dim Encoder As New IconResourceEncoder
    Dim Handle As Long
    
    Handle = LoadImage(0, MakeResourcePath(IconFileName), IMAGE_ICON, 0, 0, LR_LOADFROMFILE)
    
    Dim Actual As Boolean
    Actual = Encoder.Encode(Handle, "Test")
    
    Assert.That Actual, Iz.True
End Sub

Public Sub GetEncodedResource_WhenEncodedFromPictureHandle_ReturnsKeyAndValue()
    Dim Encoder As New IconResourceEncoder
    Dim Handle As Long
    
    Handle = LoadImage(0, MakeResourcePath(IconFileName), IMAGE_ICON, 0, 0, LR_LOADFROMFILE)
    Encoder.Encode Handle, "Test"

    Dim Key As ResourceKey
    Dim Value() As Byte
    Assert.That Encoder.GetEncodedResource(Key, Value), Iz.True
    
    Assert.That Key.ResourceName, Iz.EqualTo("Test")
    Assert.That Key.ResourceType, Iz.EqualTo(ResourceTypes.IconResource)
    Assert.That Value(0), Iz.EqualTo(40)    ' just check for correct bitmap data header size.
End Sub

