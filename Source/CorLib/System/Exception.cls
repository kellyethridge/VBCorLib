VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: Exception
'

''
' This is the base class from which all exception derive their interface.
'
' @remarks This class is the base class for all exceptions. When an
' error occurs, either the system or the currently executing application reports
' it by throwing an exception containing information about the error. Once thrown,
' an exception is handled by the application or by the default exception handler.
'
' There are two methods of raising an error:
' Using the classical VB approach <b>Err.Raise</b>. This method does not place an
' exception object in holding to be caught. In order to catch the error using
' the <b>Catch</b> method, call <b>Catch</b> passing in an exception variable and also the
' <b>Err</b> object. This will create an <b>Exception</b> object with the <b>Err</b> information.
' Use this method if an error might be raised by code that does not use the <b>Throw</b> method.
'
' Using the VBCorLib approach of <b>Throw</b> {exception object}. This method does
' place the exception in holding to be caught. It then raises an error using
' the standard <b>Err.Raise</b> method giving the information within the exception
' object. This allows for the error to be caught by error traps not utilzing the
' <b>Throw/Catch</b> style. To catch the exception using <b>Catch</b>, an error trap is still
' required. Once in the trap call <b>Catch</b> passing in an exception variable that is
' set to the held exception object. <b>Catch</b> will return True if an exception was
' caught, False otherwise.
'
' <h4>Example:</h4>
' This example demonstrates catching an exception and determining the type
' of exception caught. The <b>FileNotFoundException</b> type is tested for
' to show how exception types can be determined.
'
'
' <pre>
' Private Sub Form_Load()
'     Dim fs As FileStream
'
'     On Error GoTo CatchIt
'     Set fs = NewFileStream("missing.txt", OpenExisting)
'
' CatchIt:
'     Dim ex As Exception
'     If Catch(ex) Then
'         If TypeOf ex Is FileNotFoundException Then
'             Call Console.WriteLine("FileNotFoundException Handler: {0}", ex.ToString)
'         Else
'             Call Console.WriteLine("Generic Exception Handler: {0}", ex.ToString)
'         End If
'     End If
' End Sub
' </pre>
'
Option Explicit
Implements IObject

' Default Values
Private Const DEF_HRESULT As Long = COR_E_EXCEPTION

' Property names used with the PropertyBag
Private Const PROP_BASE As String = "Base"

' The ExceptionBase class handles everything this class will perform.
Private mBase As ExceptionBase


''
' Returns a key/value collection used to contain user-defined specific information about the exception.
'
' @return A key/value collection for user-defined information.
'
Public Property Get Data() As IDictionary
    Set Data = mBase.Data
End Property


''
' Gets a link to a help file associated with the exception.
'
' @return The Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Get HelpLink() As String
    HelpLink = mBase.HelpLink
End Property

''
' Sets a link to a help file associated with the exception.
'
' @param RHS Set the Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Let HelpLink(ByVal RHS As String)
    mBase.HelpLink = RHS
End Property

''
' Gets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @return The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned.
' @see Exception
'
Public Property Get HResult() As Long
    HResult = mBase.HResult
End Property

''
' Sets the HRESULT, a coded numerical value that is assigned to a specific exception.
'
' @param RHS The value of the associated HResult.
' @remarks An HResult is associated with an error result code. This allows for VB specific
' error codes to be returned. This is the same as <b>Err.Number</b>.
' @see Exception
'
Public Property Let HResult(ByVal RHS As Long)
    mBase.HResult = RHS
End Property

''
' Gets a description of the source of the exception.
'
' @return A description of the source of the exception.
' @remarks The source of an exception generally will contain the name of
' the function being called when the exception was thrown. This is to help
' narrow down exactly where the exception had occurred.
'
Public Property Get Source() As String
    Source = mBase.Source
End Property

''
' Sets a description of the source of the exception.
'
' @param RHS A description of the source of the exception.
' @remarks The source of an exception generally will contain the name of
' the function being called when the exception was thrown. This is to help
' narrow down exactly where the exception had occurred.
'
Public Property Let Source(ByVal RHS As String)
    mBase.Source = RHS
End Property

''
' Gets the error message associated with the Subclass exception.
'
' @return A custom message set by the Subclass, or a default
' message of "An Error has occurred."
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewException</b>.
'
Public Property Get Message() As String
    Message = mBase.Message
End Property

''
' Gets the exception that caused the Subclassed exception to be thrown.
'
' @return The inner exception that caused the current exception to be thrown.
' @remarks when an exception is thrown and that exception causes another
' exception to be thrown, then the <b>InnerException</b> of the new <b>Exception</b>
' object should contain the exception that caused it to be thrown.
'
Public Property Get InnerException() As Exception
    Set InnerException = mBase.InnerException
End Property

''
' Gets the original exception that caused the chain of exceptions to occur.
'
' @return The <b>Exception</b> that caused the chain of exceptions to occur.
' @remarks If exceptions set their <b>InnerException</b> to a previously thrown
' exception, then a chain of exceptions can be created. Using this function will
' traverse that chain of exceptions until the original exception is reached. That
' exception with then be returned to the caller.
' <p>When an <b>InnerException</b> of Nothing is reached, then then the exception object is returned
' as the base exception because it did not have an inner exception, so it is assumed that
' the exception object is the last in the chain and therefore the cause of the
' chain of exceptions being iterated.
'
Public Function GetBaseException() As Exception
    Set GetBaseException = mBase.GetBaseException
End Function

''
' Returns the exception message prepended with the type name of the Subclass Exception.
'
' @return A formatted message containing the original message and possible type of exception.
' @remarks A general format might look like this:<br>
' VBCorLib.SystemException: An Error has occurred.
' <p>A listing of all inner exceptions will be included in the return value.</p>
'
Public Function ToString() As String
    ToString = mBase.ToString(Message)
End Function

''
' Provides a basic implementation of the Equals function of the <b>IObject</b> interface.
'
' @param Value The value to determine if is the same object instance as the exception.
' @return Returns True if Value is the same instance as the exception object, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = mBase.Equals(Value)
End Function

''
' Provides a basic implementation of the GetHashcode function of the <b>IObject</b> interface.
'
' @return A 32-bit value used to help identify the exception object.
' @remarks The value generated is not unique across all hashcodes. Additional
' measures must be taken to find a unique value that happens to have the same
' hashcode as the exception object.
'
Public Function GetHashCode() As Long
    GetHashCode = mBase.GetHashCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Message As String, ByRef InnerException As Exception)
    Set mBase = Cor.NewExceptionBase(Me, App.Title, Message, InnerException, DEF_HRESULT)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Call Init(vbNullString, Nothing)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Set mBase = PropBag.ReadProperty(PROP_BASE)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(PROP_BASE, mBase)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
