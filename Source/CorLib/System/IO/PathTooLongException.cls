VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathTooLongException"
Attribute VB_GlobalNameSpace = False
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
' Module: PathTooLongException
'

''
' An exception that is caused by a path being longer than Windows can handle.
'
' @remarks The maximum length for a path in Windows is 260 characters.
'
' @see Constructors
' @see ExceptionMethods
' @see Exception
'
Option Explicit
Implements IObject
Implements Exception
Implements SystemException
Implements IOException

Private Const PropBase As String = "Base"

Private mBase As ExceptionBase


''
' Returns a key/value collection used to contain user-defined specific information about the exception.
'
' @return A key/value collection for user-defined information.
' @remarks This is a public property that is used by the exception
' classes to allow user-defined information to be included with the exception.
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
' @param Value Set the Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Let HelpLink(ByRef Value As String)
    mBase.HelpLink = Value
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
    HResult = COR_E_PATHTOOLONG
End Property

''
' Returns the VB error number associated with this exception instance.
'
' @return A VB error that matches the number return by <b>Err.Number</b>.
'
Public Property Get ErrorNumber() As Long
    ErrorNumber = corBadNameOrNumber
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
' @param Value A description of the source of the exception.
' @remarks The source of an exception generally will contain the name of
' the function being called when the exception was thrown. This is to help
' narrow down exactly where the exception had occurred.
'
Public Property Let Source(ByRef Value As String)
    mBase.Source = Value
End Property

''
' Gets the error message associated with the Subclass exception.
'
' @return A custom message set by the Subclass, or a default
' message of "An Error has occurred."
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewSystemException</b>.
'
Public Property Get Message() As String
    Message = GetExceptionMessage(mBase, IO_PathTooLong)
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
    Equals = MyBase.Equals(Me, Value)
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
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Set mBase = Cor.NewExceptionBase(Me, App)
End Sub

Friend Sub Init(ByRef Message As String, ByVal InnerException As Exception)
    mBase.Message = Message
    Set mBase.InnerException = InnerException
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Set mBase = PropBag.ReadProperty(PropBase, Nothing)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PropBase, mBase
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Exception
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Exception_Data() As IDictionary
    Set Exception_Data = Data
End Property

Private Function Exception_Equals(Value As Variant) As Boolean
    Exception_Equals = Equals(Value)
End Function

Private Property Get Exception_ErrorNumber() As Long
    Exception_ErrorNumber = ErrorNumber
End Property

Private Function Exception_GetBaseException() As Exception
    Set Exception_GetBaseException = GetBaseException
End Function

Private Function Exception_GetHashCode() As Long
    Exception_GetHashCode = GetHashCode
End Function

Private Property Let Exception_HelpLink(ByRef Value As String)
    HelpLink = Value
End Property

Private Property Get Exception_HelpLink() As String
    Exception_HelpLink = HelpLink
End Property

Private Property Get Exception_HResult() As Long
    Exception_HResult = HResult
End Property

Private Property Get Exception_InnerException() As Exception
    Set Exception_InnerException = InnerException
End Property

Private Property Get Exception_Message() As String
    Exception_Message = Message
End Property

Private Property Let Exception_Source(ByRef Value As String)
    Source = Value
End Property

Private Property Get Exception_Source() As String
    Exception_Source = Source
End Property

Private Function Exception_ToString() As String
    Exception_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IOException
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IOException_Data() As IDictionary
    Set IOException_Data = Data
End Property

Private Function IOException_Equals(Value As Variant) As Boolean
    IOException_Equals = Equals(Value)
End Function

Private Property Get IOException_ErrorNumber() As Long
    IOException_ErrorNumber = ErrorNumber
End Property

Private Function IOException_GetBaseException() As Exception
    Set IOException_GetBaseException = GetBaseException
End Function

Private Function IOException_GetHashCode() As Long
    IOException_GetHashCode = GetHashCode
End Function

Private Property Let IOException_HelpLink(ByRef Value As String)
    HelpLink = Value
End Property

Private Property Get IOException_HelpLink() As String
    IOException_HelpLink = HelpLink
End Property

Private Property Get IOException_HResult() As Long
    IOException_HResult = HResult
End Property

Private Property Get IOException_InnerException() As Exception
    Set IOException_InnerException = InnerException
End Property

Private Property Get IOException_Message() As String
    IOException_Message = Message
End Property

Private Property Let IOException_Source(ByRef Value As String)
    Source = Value
End Property

Private Property Get IOException_Source() As String
    IOException_Source = Source
End Property

Private Function IOException_ToString() As String
    IOException_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SystemException
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get SystemException_Data() As IDictionary
    Set SystemException_Data = Data
End Property

Private Function SystemException_Equals(Value As Variant) As Boolean
    SystemException_Equals = Equals(Value)
End Function

Private Property Get SystemException_ErrorNumber() As Long
    SystemException_ErrorNumber = ErrorNumber
End Property

Private Function SystemException_GetBaseException() As Exception
    Set SystemException_GetBaseException = GetBaseException
End Function

Private Function SystemException_GetHashCode() As Long
    SystemException_GetHashCode = GetHashCode
End Function

Private Property Let SystemException_HelpLink(ByRef Value As String)
    HelpLink = Value
End Property

Private Property Get SystemException_HelpLink() As String
    SystemException_HelpLink = HelpLink
End Property

Private Property Get SystemException_HResult() As Long
    SystemException_HResult = HResult
End Property

Private Property Get SystemException_InnerException() As Exception
    Set SystemException_InnerException = InnerException
End Property

Private Property Get SystemException_Message() As String
    SystemException_Message = Message
End Property

Private Property Let SystemException_Source(ByRef Value As String)
    Source = Value
End Property

Private Property Get SystemException_Source() As String
    SystemException_Source = Source
End Property

Private Function SystemException_ToString() As String
    SystemException_ToString = ToString
End Function




