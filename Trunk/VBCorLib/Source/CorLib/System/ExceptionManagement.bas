Attribute VB_Name = "ExceptionManagement"
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: modExceptionMethods
'

''
' Provides the mechanisms to Throw and Catch exceptions in the system.
'
Option Explicit

Private mException As Exception


Public Sub ClearException()
    Set mException = Nothing
End Sub

Public Function PeekException() As Exception
    Set PeekException = mException
End Function

Public Function TakeException() As Exception
    Set TakeException = mException
    Set mException = Nothing
End Function

Public Sub Throw(Optional ByVal Ex As Object)
    If Not Ex Is Nothing Then
        If TypeOf Ex Is Exception Then
            Set mException = Ex
        ElseIf TypeOf Ex Is ErrObject Then
            Dim ErrObj As ErrObject
            Set ErrObj = Ex
            Set mException = CreateException(ErrObj.Description, ErrObj.Number, ErrObj.Source, ErrObj.HelpFile)
        Else
            Set mException = Cor.NewSystemException("Invalid Throw argument. Must be an Exception type or ErrObject.")
        End If
    End If
    
    If Not mException Is Nothing Then
        Err.Raise mException.ErrorNumber, mException.Source, mException.Message
    End If
End Sub

Public Function Catch(ByRef Ex As Exception, Optional ByVal Err As ErrObject) As Boolean
    If Not mException Is Nothing Then
        Set Ex = mException
        Set mException = Nothing
        Catch = True
    ElseIf Not Err Is Nothing Then
        If Err.Number Then
            Set Ex = CreateException(Err.Description, Err.Number, Err.Source, Err.HelpFile)
            Err.Clear
            Catch = True
        End If
    End If
    VBA.Err.Clear
End Function

Public Function CatchArgument(ByRef Ex As ArgumentException) As Boolean
    If TypeOf PeekException Is ArgumentException Then
        Set Ex = TakeException
        CatchArgument = True
    End If
End Function

Public Function CatchArgumentNull(ByRef Ex As ArgumentNullException) As Boolean
    If TypeOf PeekException Is ArgumentNullException Then
        Set Ex = TakeException
        CatchArgumentNull = True
    End If
End Function

Public Function CatchArgumentOutOfRange(ByRef Ex As ArgumentOutOfRangeException) As Boolean
    If TypeOf PeekException Is ArgumentOutOfRangeException Then
        Set Ex = TakeException
        CatchArgumentOutOfRange = True
    End If
End Function

Public Function CatchDirectoryNotFound(ByRef Ex As DirectoryNotFoundException) As Boolean
    If TypeOf PeekException Is DirectoryNotFoundException Then
        Set Ex = TakeException
        CatchDirectoryNotFound = True
    End If
End Function

Public Function CatchFileNotFound(ByRef Ex As FileNotFoundException) As Boolean
    If TypeOf PeekException Is FileNotFoundException Then
        Set Ex = TakeException
        CatchFileNotFound = True
    End If
End Function

Public Function CatchPathTooLong(ByRef Ex As PathTooLongException) As Boolean
    If TypeOf PeekException Is PathTooLongException Then
        Set Ex = TakeException
        CatchPathTooLong = True
    End If
End Function

Public Function GetExceptionMessage(ByVal Base As ExceptionBase, ByVal DefaultMessageKey As ResourceStringKey, ParamArray Args() As Variant) As String
    Dim Message As String
    Message = Base.Message
    
    If CorString.IsNull(Message) Then
        Message = Environment.GetResourceString(DefaultMessageKey)
        
        If UBound(Args) >= 0 Then
            Dim Arguments() As Variant
            If InIDE Then
                Helper.Swap4 ByVal ArrPtr(Arguments), ByVal Helper.DerefEBP(20)
            Else
                Helper.Swap4 ByVal ArrPtr(Arguments), ByVal Helper.DerefEBP(16)
            End If
            Message = CorString.FormatArray(Message, Arguments)
        End If
    End If
    
    GetExceptionMessage = Message
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateException(ByRef Message As String, ByVal ErrorNumber As Long, ByRef Source As String, ByRef HelpLink As String) As Exception
    Set CreateException = Cor.NewException(Message, ErrorNumber)
    CreateException.Source = Source
    CreateException.HelpLink = HelpLink
End Function



