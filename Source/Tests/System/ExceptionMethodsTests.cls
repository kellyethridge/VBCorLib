VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionMethodsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ExceptionMethodsTests
'
Option Explicit
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub Setup()
    ClearException
End Sub

Public Sub Throw_WithNothing_DoesNothing()
    Throw
End Sub

Public Sub Throw_WithException_RaisesError()
    Dim Ex As Exception
    Set Ex = Cor.NewException("Msg", 10)
    Ex.Source = "Src"
    
    Assert.Throws 10, "Src", "Msg"
    
    Throw Ex
End Sub

Public Sub Catch_WhenExceptionThrown_ReturnsThrownException()
    On Error Resume Next
    Dim Ex As New Exception
    Throw Ex
    On Error GoTo 0
    
    Dim ActualEx As Exception
    Dim Actual As Boolean
    Actual = Catch(ActualEx)
    
    Assert.That Actual, Iz.True
    Assert.That ActualEx, Iz.SameAs(Ex)
    Assert.Pass
End Sub

Public Sub Catch_WithNoException_ReturnsNothing()
    Dim ActualEx As Exception
    Dim Actual As Boolean
    
    Actual = Catch(ActualEx)
    
    Assert.That Actual, Iz.False
    Assert.That ActualEx, Iz.Nothing
End Sub

Public Sub Throw_WithErrObject_RaisesError()
    Err.Number = 20
    Err.Description = "Mine"
    Err.Source = "Source"
    
    Assert.Throws 20, "Source", "Mine"
    
    Throw Err
End Sub

Public Sub Catch_WithThrownErrObject_ReturnsCorrispondingException()
    On Error Resume Next
    Err.Number = 20
    Err.Description = "Mine"
    Err.Source = "Source"
    Throw Err
    On Error GoTo 0
    
    Dim Actual As Boolean
    Dim ActualEx As Exception
    Actual = Catch(ActualEx)
    
    Assert.That Actual, Iz.True
    Assert.That ActualEx.ErrorNumber, Iz.EqualTo(20)
    Assert.That ActualEx.Message, Iz.EqualTo("Mine")
    Assert.That ActualEx.Source, Iz.EqualTo("Source")
    Assert.Pass
End Sub

Public Sub Catch_FromErrObject_ReturnsCaughtException()
    Err.Number = 20
    Err.Description = "Mine"
    Err.Source = "Source"
    
    Dim Actual As Boolean
    Dim ActualEx As Exception
    Actual = Catch(ActualEx, Err)
    
    Assert.That Actual, Iz.True
    Assert.That ActualEx.ErrorNumber, Iz.EqualTo(20)
    Assert.That ActualEx.Message, Iz.EqualTo("Mine")
    Assert.That ActualEx.Source, Iz.EqualTo("Source")
End Sub

Public Sub PeekException_WithThrownException_ReturnsException()
    On Error Resume Next
    Dim Ex As New ArgumentException
    Throw Ex
    On Error GoTo 0
    
    Dim Actual As Exception
    Set Actual = PeekException
    
    Assert.That Actual, Iz.SameAs(Ex)
    Assert.Pass
End Sub

Public Sub PeekException_WhenCalled_DoesNotClearException()
    On Error Resume Next
    Throw New Exception
    On Error GoTo 0
    
    PeekException
    
    Assert.That PeekException, Iz.Not.Nothing
    Assert.Pass
End Sub

Public Sub TakeException_WhenCalled_ReturnsAndClearsException()
    On Error Resume Next
    Dim Ex As New Exception
    Throw Ex
    On Error GoTo 0
    
    Dim Actual As Exception
    Set Actual = TakeException
    
    Assert.That Actual, Iz.SameAs(Ex)
    Assert.That PeekException, Iz.Nothing
    Assert.Pass
End Sub













