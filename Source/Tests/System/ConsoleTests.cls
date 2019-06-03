VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsoleTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConsoleTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System", "Console"
End Sub

Public Sub WriteValue_WithText_OutputsValueToOut()
    Dim Output As New StringWriter
    Console.SetOut Output
    
    Console.WriteValue "Hi"
    
    Assert.That Output.ToString, Iz.EqualTo("Hi")
End Sub

Public Sub WriteValue_WithArguments_OutputsValueToOut()
    Dim Output As New StringWriter
    Console.SetOut Output
    
    Console.WriteValue "Hello, {0}", "Kelly"
    
    Assert.That Output.ToString, Iz.EqualTo("Hello, Kelly")
End Sub

Public Sub WriteLine_WithNoText_OutputsNewLine()
    Dim Output As New StringWriter
    Console.SetOut Output
    
    Console.WriteLine
    
    Assert.That Output.ToString, Iz.EqualTo(vbCrLf)
End Sub

Public Sub WriteLine_WithText_OutputsValueWithNewLineAppendedToOut()
    Dim Output As New StringWriter
    Console.SetOut Output
    
    Console.WriteLine "Hi"
    
    Assert.That Output.ToString, Iz.EqualTo("Hi" & vbCrLf)
End Sub

Public Sub WriteLine_WithArguments_OutputsValueWithNewLineAppendedToOut()
    Dim Output As New StringWriter
    Console.SetOut Output

    Console.WriteLine "Hello, {0}", "Kelly"
    
    Assert.That Output.ToString, Iz.EqualTo("Hello, Kelly" & vbCrLf)
End Sub
