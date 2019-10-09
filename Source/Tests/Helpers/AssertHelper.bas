Attribute VB_Name = "AssertHelper"
Option Explicit

Public SkipUnsupportedTimeZone As Boolean

Public Sub AssertKeySizes(ByVal Actual As KeySizes, ByVal ExpectedMin As Long, ByVal ExpectedMax As Long, ByVal ExpectedSkip As Long)
    Assert.That Actual.MinSize, Iz.EqualTo(ExpectedMin), "Wrong MinSize"
    Assert.That Actual.MaxSize, Iz.EqualTo(ExpectedMax), "Wrong MaxSize"
    Assert.That Actual.SkipSize, Iz.EqualTo(ExpectedSkip), "Wrong SkipSize"
End Sub

Public Sub AssertPacificTimeZone()
    Dim Info As TIME_ZONE_INFORMATION
    Dim Result As Long
    
    Result = GetTimeZoneInformation(Info)
    If Result = TIME_ZONE_ID_INVALID Then
        Assert.Fail "Could not discover time zone information"
    End If
    
    Dim StandardName As String
    StandardName = SysAllocString(VarPtr(Info.StandardName(0)))
    
    If StandardName <> "Pacific Standard Time" Then
        If SkipUnsupportedTimeZone Then
            Assert.Pass "Time zone specific test was skipped."
        Else
            Assert.Ignore "Test only works for Pacific time zone."
        End If
    End If
End Sub

Public Sub AssertNoException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Assert.That Catch(Ex, Err), Iz.False, "An exception is not expected to be thrown."
End Sub

Public Function AssertArgumentException(ByVal Err As ErrObject, Optional ByRef ParamName As String) As ArgumentException
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)

    If TypeName(Ex) <> "ArgumentException" Then
        WrongException "ArgumentException", Ex
    End If
    
    If Len(ParamName) > 0 Then
        Dim ArgEx As ArgumentException
        Set ArgEx = Ex
        Assert.That ArgEx.ParamName, Iz.EqualTo(ParamName), "Wrong parameter name given."
    End If
    
    Set AssertArgumentException = Ex
End Function

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

Public Sub AssertFileNotFoundException(ByVal Err As ErrObject, Optional ByVal FileName As String)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is FileNotFoundException Then
        WrongException "FileNotFoundException", Ex
    End If
    
    If Len(FileName) > 0 Then
        Dim Fex As FileNotFoundException
        Set Fex = Ex
        Assert.That Fex.FileName, Iz.EqualTo(FileName)
    End If
End Sub

Public Sub AssertIOException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is IOException Then
        WrongException "IOException", Ex
    End If
End Sub

Public Sub AssertCryptographicException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is CryptographicException Then
        WrongException "CryptographicException", Ex
    End If
End Sub

Public Sub AssertXmlSyntaxException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is XmlSyntaxException Then
        WrongException "XmlSyntaxException", Ex
    End If
End Sub

Public Sub AssertUnauthorizedAccessException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is UnauthorizedAccessException Then
        WrongException "UnauthorizedAccessException", Ex
    End If
End Sub

Public Function AssertEncoderFallbackException(ByVal Err As ErrObject) As EncoderFallbackException
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is EncoderFallbackException Then
        WrongException "EncoderFallbackException", Ex
    End If

    Set AssertEncoderFallbackException = Ex
End Function

Public Function AssertDecoderFallbackException(ByVal Err As ErrObject) As DecoderFallbackException
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is DecoderFallbackException Then
        WrongException "DecoderFallbackException", Ex
    End If
    
    Set AssertDecoderFallbackException = Ex
End Function

Public Sub AssertArithmeticException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is ArithmeticException Then
        WrongException "ArithmeticException", Ex
    End If
End Sub

Public Sub AssertDivideByZeroException(ByVal Err As ErrObject)
    Dim Ex As Exception
    Set Ex = AssertExceptionThrown(Err)
    If Not TypeOf Ex Is DivideByZeroException Then
        WrongException "DivideByZeroException", Ex
    End If
End Sub

Private Function AssertExceptionThrown(ByVal Err As ErrObject) As Exception
    If Not Catch(AssertExceptionThrown, Err) Then
        Assert.Fail "An exception should be thrown."
    End If
End Function

Private Sub WrongException(ByVal Expected As String, ByVal Actual As Exception)
    Assert.Fail "Expected '" & Expected & "' but was '" & TypeName(Actual) & "'." & vbCrLf & "Message: " & Actual.Message
End Sub
