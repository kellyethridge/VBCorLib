VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriter
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "WriteValue_WithValues_WritesExpected"
            Test.Use CByte(1), NewBytes(49)
            Test.Use CInt(1), NewBytes(49)
            Test.Use CLng(1), NewBytes(49)
            Test.Use CLng(12), NewBytes(49, 50)
            Test.Use CLng(12345678), NewBytes(49, 50, 51, 52, 53, 54, 55, 56)
            Test.Use 1#, NewBytes(49)
            Test.Use 1.1, NewBytes(49, 46, 49)
            Test.Use CDec(1.1), NewBytes(49, 46, 49)
            Test.Use "Hello", NewBytes(72, 101, 108, 108, 111)
            
            Dim Obj As Int32
            Set Obj = NewInt32(10)
            Test.Use Obj, NewBytes(49, 48)
            
    End Select
End Sub

Public Sub NewStreamWriter_WithNothingSource_ThrowsArgumentNullException()
    On Error Resume Next
    
    Cor.NewStreamWriter Nothing
    
    AssertArgumentNullException Err, "Source"
End Sub

Public Sub NewStreamWriter_WithNonStreamObject_ThrowsArgumentException()
    On Error Resume Next
    
    Cor.NewStreamWriter Me
     
    AssertArgumentException Err, "Source"
End Sub

Public Sub NewStreamWriter_WithNegativeBufferSize_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewStreamWriter New MemoryStream, BufferSize:=-1
    
    AssertArgumentOutOfRangeException Err, "BufferSize"
End Sub

Public Sub NewStreamWriter_WithNoEncoding_DefaultsToUTF8Encoding()
    Dim Actual As Encoding
    
    With Cor.NewStreamWriter(New MemoryStream)
        Set Actual = .Encoding
    End With
    
    Assert.That TypeOf Actual Is UTF8Encoding, Iz.True
End Sub

Public Sub WriteValue_WithValues_WritesExpected(ByRef Value As Variant, ByRef Expected() As Byte)
    Dim Store As New MemoryStream
        
    With NewStreamWriter(Store)
        .WriteValue Value
    End With
    
    Dim Actual() As Byte
    Actual = Store.ToArray
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub WriteChars_WithNullChars_ThrowsArgumentNullException()
    Dim NullChars() As Integer
    On Error Resume Next
    
    With Cor.NewStreamWriter(New MemoryStream)
        .WriteChars NullChars
    End With
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub WriteChars_WithChars_WritesToStream()
    Dim Store   As New MemoryStream
    Dim Chars() As Integer
    Chars = NewIntegers(48, 49, 50)
    
    With Cor.NewStreamWriter(Store)
        .WriteChars Chars
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(48, 49, 50))
End Sub

Public Sub WriteChars_WithSubsetSpecified_WritesSubsetToStream()
    Dim Store   As New MemoryStream
    Dim Chars() As Integer
    Chars = NewIntegers(48, 48, 48, 49, 50, 51, 48, 48, 48)
    
    With Cor.NewStreamWriter(Store)
        .WriteChars Chars, 3, 3
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(49, 50, 51))
End Sub

Public Sub WriteValue_WithFormat_WritesToStream()
    Dim Store As New MemoryStream
    
    With Cor.NewStreamWriter(Store)
        .WriteValue "He{0}lo", "l"
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(72, 101, 108, 108, 111))
End Sub

Public Sub WriteLine_WithNoValue_WritesNewLine()
    Dim Store As New MemoryStream
    
    With Cor.NewStreamWriter(Store)
        .WriteLine
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(13, 10))
End Sub

Public Sub WriteLine_WithValue_WritesLineAndNewLine()
    Dim Store As New MemoryStream
    
    With Cor.NewStreamWriter(Store)
        .WriteLine "111"
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(49, 49, 49, 13, 10))
End Sub

Public Sub WriteLine_WriteFormat_WritesToStream()
    Dim Store As New MemoryStream
    
    With Cor.NewStreamWriter(Store)
        .WriteLine "He{0}lo", "l"
    End With
    
    Assert.That Store.ToArray, Iz.EqualTo(NewBytes(72, 101, 108, 108, 111, 13, 10))
End Sub

Public Sub NewStreamWriter_WhenAppending_AppendsToStream()
    Dim IO As New IOHelper
    Dim FileName As String
    FileName = IO.CreateFile(Data:="1")
    
    With Cor.NewStreamWriter(FileName, True)
        .WriteValue "234"
    End With
    
    Dim Actual As String
    Actual = IO.ReadFile(FileName)
    Assert.That Actual, Iz.EqualTo("1234")
End Sub

Public Sub Flush_WhenCalled_WritesToStream()
    Dim Store As New MemoryStream
    Dim Actual() As Byte
    
    With Cor.NewStreamWriter(Store)
        .WriteValue "Hi"
        .Flush
        Actual = Store.ToArray
    End With
    
    Assert.That Actual, Iz.EqualTo(NewBytes(Asc("H"), Asc("i")))
End Sub

Public Sub WriteValue_WithAutoFluseEnabled_WritesToStream()
    Dim Store As New MemoryStream
    Dim Actual() As Byte
    
    With Cor.NewStreamWriter(Store)
        .AutoFlush = True
        .WriteValue "12"
        Actual = Store.ToArray
    End With
    
    Assert.That Actual, Iz.EqualTo(NewBytes(49, 50))
End Sub

Public Sub WriteChars_WithAutoFlushEnabled_WritesToStream()
    Dim Store As New MemoryStream
    Dim Actual() As Byte
    
    With Cor.NewStreamWriter(Store)
        .AutoFlush = True
        .WriteChars NewIntegers(49, 50)
        Actual = Store.ToArray
    End With
    
    Assert.That Actual, Iz.EqualTo(NewBytes(49, 50))
End Sub

Public Sub WriteLine_WithAutoFlushEnabled_WritesToStream()
    Dim Store As New MemoryStream
    Dim Actual() As Byte
    
    With Cor.NewStreamWriter(Store)
        .AutoFlush = True
        .WriteLine
        Actual = Store.ToArray
    End With
    
    Assert.That Actual, Iz.EqualTo(NewBytes(13, 10))
End Sub

Public Sub WriteCharsLine_WhenCalled_WritesCharsAndNewLine()
    Dim Store As New MemoryStream
    Dim Actual() As Byte
    
    With Cor.NewStreamWriter(Store)
        .WriteLineChars NewChars("0", "1")
    End With
    
    Actual = Store.ToArray
    Assert.That Actual, Iz.EqualTo(NewBytes(48, 49, 13, 10))
End Sub









