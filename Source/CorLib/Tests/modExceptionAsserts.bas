Attribute VB_Name = "modExceptionAsserts"
Option Explicit

Public Sub AssertArgumentException(ByVal Err As ErrObject, Optional ByRef ParamName As String)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArgumentException Then
        Assert.Fail "Expected an ArgumentException but was '" & TypeName(Ex) & "'."
    End If
    
    If Len(ParamName) > 0 Then
        Dim ArgEx As ArgumentException
        Set ArgEx = Ex
        Assert.That ArgEx.ParamName, Iz.EqualTo(ParamName), "Wrong parameter name given."
    End If
End Sub

Public Sub AssertArgumentNullException(ByVal Err As ErrObject, ByRef ParamName As String)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArgumentNullException Then
        Assert.Fail "Expected an ArgumentNullException but was '" & TypeName(Ex) & "'."
    End If
    Dim ArgEx As ArgumentNullException
    Set ArgEx = Ex
    Assert.That ArgEx.ParamName, Iz.EqualTo(ParamName), "Wrong parameter name given."
End Sub

Public Sub AssertArgumentOutOfRangeException(ByVal Err As ErrObject, Optional ByRef ParamName As String, Optional ByRef ActualValue As Variant)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArgumentOutOfRangeException Then
        Assert.Fail "Expected an ArgumentOutOfRangeException but was " & TypeName(Ex) & "."
    End If
    
    If Len(ParamName) > 0 Then
        Dim ArgEx As ArgumentOutOfRangeException
        Set ArgEx = Ex
        Assert.That ArgEx.ParamName, Iz.EqualTo(ParamName), "Wrong parameter name given."
    End If
    
    If Not IsMissing(ActualValue) Then
        Assert.That ArgEx.ActualValue, Iz.EqualTo(ActualValue), "Wrong actual value given."
    End If
End Sub

Public Sub AssertIndexOutOfRangeException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is IndexOutOfRangeException Then
        Assert.Fail "Expected an IndexOutOfRangeException but was '" & TypeName(Ex) & "'."
    End If
End Sub


Private Function AssertExceptionThrown(ByVal Err As ErrObject) As Exception
    If Not catch(AssertExceptionThrown, Err) Then
        Assert.Fail "An exception should be thrown."
    End If
End Function
