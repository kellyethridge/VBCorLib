VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringWriterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestStringWriter
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
            Test.Use 1, "1"
            Test.Use 1234.567, "1234.567"
            Test.Use #2/2/2004 3:00:00 PM#, "2/2/2004 3:00:00 PM"
            Test.Use NewInt32(55), "55"
            
        Case "WriteLine_WithValues_WritesExpected"
            Test.Use 1, "1" & vbCrLf
            Test.Use 1234.567, "1234.567" & vbCrLf
            Test.Use #2/2/2004 3:00:00 PM#, "2/2/2004 3:00:00 PM" & vbCrLf
            Test.Use NewInt32(55), "55" & vbCrLf
        
    End Select
End Sub

Public Sub Encoding_WhenCalled_ReturnsUnicodeEncoding()
    Dim Writer As TextWriter
    Set Writer = CreateWriter
    
    Dim Actual As Encoding
    Set Actual = Writer.Encoding
    
    Assert.That TypeOf Actual Is UnicodeEncoding, Iz.True
End Sub

Public Sub FormatProvider_WhenCalled_ReturnsEnglishUS()
    Dim Writer As TextWriter
    Set Writer = CreateWriter
    
    Dim Actual As IFormatProvider
    Set Actual = Writer.FormatProvider
    
    Assert.That TypeOf Actual Is CultureInfo, Iz.True
    Dim Culture As CultureInfo
    Set Culture = Actual
    Assert.That Culture.Name, Iz.EqualTo("en-US")
End Sub

Public Sub WriteValue_WithValues_WritesExpected(ByRef Value As Variant, ByVal Expected As String)
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    Writer.WriteValue Value
    
    Assert.That Writer.ToString, Iz.EqualTo(Expected)
End Sub

Public Sub WriteValue_WithFormat_WritesFormattedValue()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    Writer.WriteValue "It is {0}", #1/1/2001#
    
    Assert.That Writer.ToString, Iz.EqualTo("It is 1/1/2001 12:00:00 AM")
End Sub

Public Sub NewStringWriter_WithStringBuilder_UsesProvidedStringBuilder()
    Dim sb As New StringBuilder
    Dim Writer As StringWriter
    Set Writer = Cor.NewStringWriter(sb)
    
    Writer.WriteValue "Hello"
    
    Assert.That sb.ToString, Iz.EqualTo("Hello")
End Sub

Public Sub WriteValue_WithFrenchCulture_WritesExpected()
    Dim Writer As StringWriter
    Set Writer = CreateWriter(Cor.NewCultureInfo("fr-FR"))
    
    Writer.WriteValue 1234.567
    
    Assert.That Writer.ToString, Iz.EqualTo("1234,567")
End Sub

Public Sub WriteValue_WithFormatAndFrenchCulture_WritesExpected()
    Dim Writer As StringWriter
    Set Writer = CreateWriter(Cor.NewCultureInfo("fr-FR"))
    
    Writer.WriteValue "Number={0}", 1234.567
    
    Assert.That Writer.ToString, Iz.EqualTo("Number=1234,567")
End Sub

Public Sub WriteLine_WithNoValue_WritesNewLine()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    Writer.WriteLine
    
    Assert.That Writer.ToString, Iz.EqualTo(vbCrLf)
End Sub

Public Sub WriteLine_WithValues_WritesExpected(ByRef Value As Variant, ByVal Expected As String)
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    Writer.WriteLine Value
    
    Assert.That Writer.ToString, Iz.EqualTo(Expected)
End Sub

Public Sub WriteLine_WithFormatAndFrenchCulture_WritesExpected()
    Dim Writer As StringWriter
    Set Writer = CreateWriter(Cor.NewCultureInfo("fr-FR"))
    
    Writer.WriteLine "Number={0}", 1234.567
    
    Assert.That Writer.ToString, Iz.EqualTo("Number=1234,567" & vbCrLf)
End Sub

Public Sub WriteLine_WithFrenchCulture_WritesExpected()
    Dim Writer As StringWriter
    Set Writer = CreateWriter(Cor.NewCultureInfo("fr-FR"))
    
    Writer.WriteLine 1234.567
    
    Assert.That Writer.ToString, Iz.EqualTo("1234,567" & vbCrLf)
End Sub

Public Sub WriteChars_WithNullChars_ThrowsArgumentNullException()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    Dim NullChars() As Integer
    
    On Error Resume Next
    Writer.WriteChars NullChars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub WriteChars_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    On Error Resume Next
    Writer.WriteChars NewChars("a"), -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub WriteChars_WithCountLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    On Error Resume Next
    Writer.WriteChars NewChars("a"), 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub WriteChars_WithIndexPlusCountInvalid_ThrowsArgumentException()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    On Error Resume Next
    Writer.WriteChars NewChars("a", "b"), 1, 2
    
    AssertArgumentException Err
End Sub

Public Sub WriteChars_WithEmptyChars_WritesNothing()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    Writer.WriteChars NewChars()
    
    Assert.That Writer.ToString, Iz.Empty
End Sub

Public Sub WriteChars_WithChars_WritesChars()
    Dim Writer As StringWriter
    Set Writer = CreateWriter
    
    Writer.WriteChars NewChars("a", "b")
    
    Assert.That Writer.ToString, Iz.EqualTo("ab")
End Sub


Private Function CreateWriter(Optional ByVal Provider As IFormatProvider) As StringWriter
    If Provider Is Nothing Then
        Set Provider = Cor.NewCultureInfo("en-US")
    End If
    
    Set CreateWriter = Cor.NewStringWriter(Provider:=Provider)
End Function






























