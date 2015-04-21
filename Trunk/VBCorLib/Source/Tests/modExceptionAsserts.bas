Attribute VB_Name = "modExceptionAsserts"
Option Explicit

Public Sub AssertArgumentException(ByVal Err As ErrObject, Optional ByRef ParamName As String)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArgumentException Then
        WrongException "ArgumentException", Ex
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
        WrongException "ArgumentNullException", Ex
    End If
    Dim ArgEx As ArgumentNullException
    Set ArgEx = Ex
    Assert.That ArgEx.ParamName, Iz.EqualTo(ParamName), "Wrong parameter name given."
End Sub

Public Sub AssertArgumentOutOfRangeException(ByVal Err As ErrObject, Optional ByRef ParamName As String, Optional ByRef ActualValue As Variant)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArgumentOutOfRangeException Then
        WrongException "ArgumentOutOfRangeException", Ex
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
        WrongException "IndexOutOfRangeException", Ex
    End If
End Sub

Public Sub AssertArrayTypeMismatchException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArrayTypeMismatchException Then
        WrongException "ArrayTypeMismatchException", Ex
    End If
End Sub

Public Sub AssertInvalidOperationException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is InvalidOperationException Then
        WrongException "InvalidOperationException", Ex
    End If
End Sub

Public Sub AssertRankException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is RankException Then
        WrongException "RankException", Ex
    End If
End Sub

Public Sub AssertFormatException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is FormatException Then
        WrongException "FormatException", Ex
    End If
End Sub

Public Sub AssertOverflowException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is OverflowException Then
        WrongException "OverflowException", Ex
    End If
End Sub

Public Sub AssertEndOfStreamException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is EndOfStreamException Then
        WrongException "EndOfStreamException", Ex
    End If
End Sub

Public Sub AssertNotSupportedException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is NotSupportedException Then
        WrongException "NotSupportedException", Ex
    End If
End Sub

Public Sub AssertObjectDisposedException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ObjectDisposedException Then
        WrongException "ObjectDisposedException", Ex
    End If
End Sub

Public Sub AssertInvalidCastException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is InvalidCastException Then
        WrongException "InvalidCastException", Ex
    End If
End Sub

Public Sub AssertFileNotFoundException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is FileNotFoundException Then
        WrongException "FileNotFoundException", Ex
    End If
End Sub

Public Sub AssertIOException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is IOException Then
        WrongException "IOException", Ex
    End If
End Sub

Private Function AssertExceptionThrown(ByVal Err As ErrObject) As Exception
    If Not Catch(AssertExceptionThrown, Err) Then
        Assert.Fail "An exception should be thrown."
    End If
End Function

Private Sub WrongException(ByVal Expected As String, ByVal Actual As Exception)
    Assert.Fail "Expected '" & Expected & "' but was '" & TypeName(Actual) & "'."
End Sub
