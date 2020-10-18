VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionMethods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: ExceptionMethods
'

''
' Defines public exception management functions <b>Throw</b> and <b>Catch</b>.
'
' @remarks
' Error processing can be a fundamental process in creating robust applications
' in Visual Basic. <b>VBCorLib</b> provides a wide range of exception classes
' that can be used to capture details for an error that may occur.
' <p>By using <b>Throw</b> and <b>Catch</b> to handle exception objects, a
' robust error handling mechanism can be created.
'
' @see Exception
'
Option Explicit

''
' Catch provides a mechanism for retrieving an exception object that was
' set using the Throw function.
'
' @param ex <i>ex</i> is set to the exception object that is caught.
' @param Err Used when the exception was not thrown by <b>VBCorLib</b> directly.
' @return Indicates if an exception was actually caught.
' @remarks Returns True if an exception has been caught, setting the <i>ex</i> parameter
' to the exception object. By retrieving an exception object, the error that
' occurred can be determined along with any specific details included. In order to
' catch an exception, error trapping must be used as normal. During error trapping
' procedures, a call to Catch will attempt to retrieve an exception, returning an
' indication on whether an exception was caught.
' <p>If Throw was not used to raise an error, then the ErrObject must be passed into
' the function to retrieve the error information. This is because each loaded
' module (dll, exe, ...) will have its own ErrObject, so <b>VBCorLib</b> has no
' way of retrieving the error information for modules other than itself. If errors
' outside of <b>VBCorLib</b> are created using Throw, then the ErrObject does not
' need to be passed into the Catch function.</p>
' <p>If the source of the error will always be from <b>VBCorLib</b> then the ErrObject
' does not need to be supplied. If the source is not <b>VBCorLib</b> or is unknown, then
' pass in the ErrObject. Passing in the ErrObject will not affect catching exceptions
' thrown by <b>VBCorLib</b>.</p>
' <p>To clear an existing exception, call Catch and pass in Nothing as
' the exception variable.</p>
'
' <h4>Example</h4>
' This shows how to catch an exception thrown within the same function.
' <pre>
'     On Error Goto errTrap
'     Throw NewFileNotFoundException("Could not find your file.")
'
'    ''... code
'
' errTrap:
'     Dim ex As Exception
'
'     If Catch(ex) Then
'         If TypeOf ex Is FileNotFoundException Then
'             MsgBox "You will need to enter another filename."
'         Else
'             MsgBox "Error: " & ex.Message
'         End If
'     End If
' </pre>
' Once an exception has been caught, it can be inspected to determine the
' appropriate course of action.
' <p>The error trap section is a typical implementation of catching a <b>VBCorLib</b> exception.</p>
' <p>If an error is raised using the Throw function, then the exception object will
' remain ready to be caught until Catch is called. This is similar to using
' On Error Resume Next and not clearing any errors that may occur. The ErrObject will
' contain any error information until some event clears it. This can lead to problems
' later on in the code when an exception is checked using Catch, and an older exception
' is still set. The exception may need to be cleared to prevent unwanted exceptions
' from being caught.</p>
'
' @see Constructors
' @see Exception
'
Public Function Catch(ByRef Ex As Exception, Optional ByVal Err As ErrObject) As Boolean
    Catch = ExceptionManagement.Catch(Ex, Err)
End Function

''
' Raises an error using the supplied exception object. The exception is
' then set and ready to be caught using Catch.
'
' @param Ex The exception object used to raise an error. This will accept any
' object that implements the <b>Exception</b> interface or an <b>ErrObject</b>.
' All exception classes in <b>VBCorLib</b> implement <b>Exception</b>.
' @remarks <b>VBCorLib</b> provides a wide range of exception objects
' that can be used to define specific errors that occur during the
' execution of an application. In order to use the exception objects
' through <b>VBCorLib</b>, they must first be Thrown. Throwing an
' exception object is similar to raising an error using Visual Basic's
' standard method of Err.Raise.
' <p>To throw an exception, it must be passed into the Throw function.
' This will store the exception object so it can be retrieved (Caught)
' during error trapping.</p>
' <pre>
' '' Throwing an exception directly is the simplest method.
' Throw New Exception
' </pre>
' This creates a new <b>Exception</b> object and passes it directly into
' the Throw function. The function then sets the object to be caught and
' raises an error using the HResult, Source, and Message properties. These
' properties corrispond to the Err.Number, Err.Source, and Err.Description
' parameters used when raising an error.
' <p>Exceptions that are to be thrown should be constructed with meaningful
' information before being passed into the Throw function.</p>
' <p>An Exception that occurs but is not caught using Catch can be rethrown
' by calling Throw without any parameters. This causes the Throw function
' to check if a pending Exception already exists, and if so, rethrow it.
' This does not clear an exception.</p>
'
' @see Constructors
' @see Exception
'
Public Sub Throw(Optional ByVal Ex As Object)
    ExceptionManagement.Throw Ex
End Sub

''
' Clears any exception that may not have been caught yet.
'
' @remarks Sometimes an exception can be thrown, but ignored and not caught.
' This could cause problems later in code if the <b>Catch</b> method is called
' and an exception has not been thrown. The previous exception may still remain
' ready to be caught even though the error trapping has passed for that
' previous exception.
'
' @see Exception
'
Public Sub ClearException()
    ExceptionManagement.ClearException
End Sub

Public Function PeekException() As Exception
    Set PeekException = ExceptionManagement.PeekException
End Function

Public Function TakeException() As Exception
    Set TakeException = ExceptionManagement.TakeException
End Function

Public Function CatchArgument(ByRef Ex As ArgumentException) As Boolean
    CatchArgument = ExceptionManagement.CatchArgument(Ex)
End Function

Public Function CatchArgumentNull(ByRef Ex As ArgumentNullException) As Boolean
    CatchArgumentNull = ExceptionManagement.CatchArgumentNull(Ex)
End Function

Public Function CatchArgumentOutOfRange(ByRef Ex As ArgumentOutOfRangeException) As Boolean
    CatchArgumentOutOfRange = ExceptionManagement.CatchArgumentOutOfRange(Ex)
End Function

Public Function CatchDirectoryNotFound(ByRef Ex As DirectoryNotFoundException) As Boolean
    CatchDirectoryNotFound = ExceptionManagement.CatchDirectoryNotFound(Ex)
End Function

Public Function CatchFileNotFound(ByRef Ex As FileNotFoundException) As Boolean
    CatchFileNotFound = ExceptionManagement.CatchFileNotFound(Ex)
End Function

Public Function CatchPathTooLong(ByRef Ex As PathTooLongException) As Boolean
    CatchPathTooLong = ExceptionManagement.CatchPathTooLong(Ex)
End Function













