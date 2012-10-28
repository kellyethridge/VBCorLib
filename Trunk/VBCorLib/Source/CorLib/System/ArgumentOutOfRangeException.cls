VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentOutOfRangeException"
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
'    Module: ArgumentOutOfRangeException
'

''
' The exception that is thrown when the value of an argument is outside the
' allowable range of values as defined by the invoked method.
'
' @remarks
' <p><b>ArgumentOutOfRangeException</b> is thrown when an argument is not Null and
' is outside of the allowable range of values.</p>
' <p>An <b>ArgumentOutOfRangeException</b> works the same as an <b>ArgumentException</b>. It
' is used to distinguish between other argument exceptions and those that are out of a valid range.
'
' @see Constructors
' @see ExceptionMethods
' @see Exception
' @see ArgumentException
' @see SystemException
'
Option Explicit
Implements IObject
Implements Exception
Implements SystemException
Implements ArgumentException

Private Const PROP_PARAMNAME    As String = "ParamName"
Private Const PROP_ACTUALVALUE  As String = "ActualValue"
Private Const PROP_BASE         As String = "Base"
Private Const DEF_PARAMNAME     As String = vbNullString
Private Const DEF_ACTUALVALUE   As Variant = Empty
Private Const DEF_HRESULT       As Long = COR_E_ARGUMENTOUTOFRANGE

' This can do it all.
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
' Returns the actual value that caused the exception to be thrown.
'
' @return A value of the parameter that caused the exception.
'
Public Property Get ActualValue() As Variant
    Call Helper.MoveVariant(ActualValue, mBase.GetValue(PROP_ACTUALVALUE, MissingVariant))
End Property


''
' Returns the parameter name that caused the exception.
'
' @return The name of the parameter that caused the exception.
' @remarks Every <b>ArgumentNullException</b> should contain the name
' of the parameter that caused the exception.
' @see Exception
'
Public Property Get ParamName() As String
    ParamName = mBase.GetValue(PROP_PARAMNAME, DEF_PARAMNAME)
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
' <p>This property can be set in the constructor <b>NewArgumentNullException</b>.
'
Public Property Get Message() As String
    Message = mBase.Message
    
    ' If there is a parameter name, add it to the message
    ' on a new line for everyone to see.
    Dim ParamName As String
    ParamName = Me.ParamName
    If Len(ParamName) > 0 Then Message = Message & vbCrLf & vbCrLf & "Parameter Name: " & ParamName
    
    ' Get a string representation of the actual value.
    Dim ActualValue As String
    ActualValue = Convert.ToString(Me.ActualValue)
    
    ' If we have an actual value, we want to append it to the
    ' end of the message on it's own line, if the message is not empty.
    If Not cString.IsNull(ActualValue) Then
        If Len(Message) > 0 Then Message = Message & vbCrLf
        Message = Message & "Actual Value: " & ActualValue
    End If
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
Friend Sub Init(ByRef Message As String, ByRef ParamName As String, ByRef ActualValue As Variant, ByRef InnerException As Exception)
    Set mBase = Cor.NewExceptionBase(Me, App.Title, Message, InnerException, DEF_HRESULT)
    If Not IsMissing(ActualValue) Then Call mBase.SetValue(PROP_ACTUALVALUE, ActualValue)
    If Len(ParamName) > 0 Then Call mBase.SetValue(PROP_PARAMNAME, ParamName)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Call Init(Environment.GetResourceString(ArgumentOutOfRange_Exception), DEF_PARAMNAME, MissingVariant, Nothing)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Set mBase = PropBag.ReadProperty(PROP_BASE, Nothing)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_BASE, mBase
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ArgumentException Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ArgumentException_Data() As IDictionary
    Set ArgumentException_Data = Data
End Property

Private Function ArgumentException_Equals(Value As Variant) As Boolean
    ArgumentException_Equals = Equals(Value)
End Function

Private Function ArgumentException_GetBaseException() As Exception
    Set ArgumentException_GetBaseException = GetBaseException
End Function

Private Function ArgumentException_GetHashCode() As Long
    ArgumentException_GetHashCode = GetHashCode
End Function

Private Property Let ArgumentException_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get ArgumentException_HelpLink() As String
    ArgumentException_HelpLink = HelpLink
End Property

Private Property Let ArgumentException_HResult(ByVal RHS As Long)
    HResult = RHS
End Property

Private Property Get ArgumentException_HResult() As Long
    ArgumentException_HResult = HResult
End Property

Private Property Get ArgumentException_InnerException() As Exception
    Set ArgumentException_InnerException = InnerException
End Property

Private Property Get ArgumentException_Message() As String
    ArgumentException_Message = Message
End Property

Private Property Get ArgumentException_ParamName() As String
    ArgumentException_ParamName = ParamName
End Property

Private Property Let ArgumentException_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get ArgumentException_Source() As String
    ArgumentException_Source = Source
End Property

Private Function ArgumentException_ToString() As String
    ArgumentException_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
'   Exception Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Exception_Data() As IDictionary
    Set Exception_Data = Data
End Property

Private Function Exception_Equals(Value As Variant) As Boolean
    Exception_Equals = Equals(Value)
End Function

Private Function Exception_GetBaseException() As Exception
    Set Exception_GetBaseException = GetBaseException
End Function

Private Function Exception_GetHashCode() As Long
    Exception_GetHashCode = GetHashCode
End Function

Private Property Let Exception_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get Exception_HelpLink() As String
    Exception_HelpLink = HelpLink
End Property

Private Property Let Exception_HResult(ByVal RHS As Long)
    HResult = RHS
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

Private Property Let Exception_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get Exception_Source() As String
    Exception_Source = Source
End Property

Private Function Exception_ToString() As String
    Exception_ToString = ToString
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SystemException Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get SystemException_Data() As IDictionary
    Set SystemException_Data = Data
End Property

Private Function SystemException_Equals(Value As Variant) As Boolean
    SystemException_Equals = Equals(Value)
End Function

Private Function SystemException_GetBaseException() As Exception
    Set SystemException_GetBaseException = GetBaseException
End Function

Private Function SystemException_GetHashCode() As Long
    SystemException_GetHashCode = GetHashCode
End Function

Private Property Let SystemException_HelpLink(ByVal RHS As String)
    HelpLink = RHS
End Property

Private Property Get SystemException_HelpLink() As String
    SystemException_HelpLink = HelpLink
End Property

Private Property Let SystemException_HResult(ByVal RHS As Long)
    HResult = RHS
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

Private Property Let SystemException_Source(ByVal RHS As String)
    Source = RHS
End Property

Private Property Get SystemException_Source() As String
    SystemException_Source = Source
End Property

Private Function SystemException_ToString() As String
    SystemException_ToString = ToString
End Function




