Attribute VB_Name = "modExceptionMethods"
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


''
' Catches an exception, if one has been thrown. Otherwise, it may create
' a new exception if the Err object contains an error.
'
' @param ex The variable used to retrieve the exception that has been thrown.
' @param Err An object used to determine if an error has been raised and
' if an exception object should be created and returned if no exception
' currently has been thrown.
' @return Returns True if an exception was indeed caught, otherwise False.
'
Public Function Catch(ByRef Ex As Exception, Optional ByVal Err As ErrObject) As Boolean
    If Not mException Is Nothing Then
        Set Ex = mException
        Set mException = Nothing
        Catch = True
    ElseIf Not Err Is Nothing Then
        If Err.Number Then
            Set Ex = Cor.NewException(Err.Description)
            Ex.HResult = Err.Number
            Ex.Source = Err.Source
            Err.Clear
            Catch = True
        End If
    End If
    VBA.Err.Clear
End Function

''
' Stores the exception locally and raises an error to notify the application
' that an exception is ready to be caught.
'
' @param ex The exception to cache to be caught.
'
Public Sub Throw(Optional ByVal Ex As Object)
    If Not Ex Is Nothing Then
        If TypeOf Ex Is Exception Then
            Set mException = Ex
        ElseIf TypeOf Ex Is ErrObject Then
            Dim ErrObj As ErrObject
            Set ErrObj = Ex
            Set mException = Cor.NewException(ErrObj.Description)
            mException.Source = ErrObj.Source
            mException.HResult = ErrObj.Number
            mException.HelpLink = ErrObj.HelpFile
        Else
            Set mException = Cor.NewSystemException("Invalid Throw argument. Must be an Exception type or ErrObject.")
        End If
    End If
    
    If Not mException Is Nothing Then
        Call Err.Raise(mException.HResult, mException.Source, mException.Message)
    End If
End Sub

''
' Clears a cached exception if one exists.
'
Public Sub ClearException()
    Set mException = Nothing
End Sub

''
' Gets a formatted message from a system error code.
'
' @param MessageID The error code to retrieve the message for.
' @return A system message representing the code, or "Unknown Error." if the code was not found.
'
Public Function GetErrorMessage(ByVal MessageID As Long) As String
    Dim Buf As String
    Dim Size As Long
    
    Buf = String$(1024, vbNullChar)
    Size = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, MessageID, 0, Buf, Len(Buf), ByVal 0&)
    If Size > 0 Then
        GetErrorMessage = Left$(Buf, Size - 2)
    Else
        GetErrorMessage = "Unknown Error."
    End If
End Function
