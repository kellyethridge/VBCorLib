VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsoleKeyInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConsoleKeyInfoTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use Cor.NewConsoleKeyInfo(32, SpaceBarKey, True, True, True), True
            Test.Use Cor.NewConsoleKeyInfo(33, SpaceBarKey, True, True, True), False
            Test.Use Cor.NewConsoleKeyInfo(32, AKey, True, True, True), False
            Test.Use Cor.NewConsoleKeyInfo(32, SpaceBarKey, False, True, True), False
            Test.Use Cor.NewConsoleKeyInfo(32, SpaceBarKey, True, False, True), False
            Test.Use Cor.NewConsoleKeyInfo(32, SpaceBarKey, True, True, False), False
            Test.Use "abc", False
            Test.Use Nothing, False
            
    End Select
End Sub

Public Sub NewConsoleKeyInfo_WhenCalled_SetsProperties()
    Dim Info As ConsoleKeyInfo
    
    Set Info = Cor.NewConsoleKeyInfo(32, SpaceBarKey, True, True, True)
    
    Assert.That Info.KeyChar, Iz.EqualTo(32), "#1"
    Assert.That Info.Key, Iz.EqualTo(ConsoleKey.SpaceBarKey), "#2"
    Assert.That Info.Modifiers And ConsoleModifiers.ShiftKey, Iz.EqualTo(ConsoleModifiers.ShiftKey), "#3"
    Assert.That Info.Modifiers And ConsoleModifiers.AltKey, Iz.EqualTo(ConsoleModifiers.AltKey), "#4"
    Assert.That Info.Modifiers And ConsoleModifiers.ControlKey, Iz.EqualTo(ConsoleModifiers.ControlKey), "#5"
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef Other As Variant, ByVal Expected As Boolean)
    Dim Info As ConsoleKeyInfo
    
    Set Info = Cor.NewConsoleKeyInfo(32, SpaceBarKey, True, True, True)
    
    Assert.That Info.Equals(Other), Iz.EqualTo(Expected)
End Sub

