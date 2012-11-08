VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kell    y Ethridge
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
'    Module: ExceptionBase
'

''
' This is a helper class to easily manage an exception type class.
'
' @remarks This class should be used as a contained object that functions
' can be delegated to and performed with. This class contains the primary
' functions and properties common to all exceptions.
' <p>If the user-derived exception class needs to handle more information than
' this base class manages, then the developer needs to implement those
' additional features directly into the derived class.</p>
'
Option Explicit

' Default Values
Private Const DEF_HRESULT           As Long = 5

' Property Names used with PropertyBag
Private Const PROP_MESSAGE          As String = "Message"
Private Const PROP_ISNULLMESSAGE    As String = "IsNullMessage"
Private Const PROP_HRESULT          As String = "HResult"
Private Const PROP_SOURCE           As String = "Source"
Private Const PROP_HELPLINK         As String = "HelpLink"
Private Const PROP_INNEREXCEPTION   As String = "InnerException"
Private Const PROP_VALUES           As String = "Values"
Private Const PROP_DATA             As String = "Data'"

' Set of values that are common to all exception classes.
Private mHResult            As Long
Private mMessage            As String
Private mSource             As String
Private mHelpLink           As String
Private mInnerException     As Exception
Private mValues             As Hashtable
Private mSubclass           As Long
Private mSubclassName       As String
Private mData               As Hashtable


''
' Returns a key/value set of data that is associated to the exception.
'
' @return a key/value set of data that is associated to the exception.
' @remarks This is a public property that is used by the exception
' classes to allow user-defined information to be included with the exception.
'
Public Property Get Data() As IDictionary
    If mData Is Nothing Then Set mData = New Hashtable
    Set Data = mData
End Property

''
' Sets a value that can be associated with the exception without
' needing to maintain the value in the Subclass exception class.
'
' @param Name The name of the value to set.
' @param Value The value to set.
' @remarks If the value already exists, it will be replaced.
'
Public Sub SetValue(ByVal Name As String, ByRef Value As Variant)
    If IsObject(Value) Then
        Set Values(Name) = Value
    Else
        Values(Name) = Value
    End If
End Sub

''
' Returns a value being maintained in the exception base class.
'
' @param Name The name of the value to retrieve.
' @param Default The value to be returned if the value doesn't exist.
'
Public Function GetValue(ByVal Name As String, Optional ByRef Default As Variant) As Variant
    ' Steal the returned value from the Hashtable.
    Call Helper.MoveVariant(GetValue, Values(Name))
    
    ' An empty value is the default return value for a key
    ' that does not exist in a Hashtable. It is also a valid
    ' value that may have been added to the Hashtable.
    ' So now we can check if the key exists. If it doesn't,
    ' then return the default value. This saves us from
    ' having to call into the Hashtable twice unless the
    ' returned value is Empty.
    If IsEmpty(GetValue) Then
        If Not Values.Contains(Name) Then Call VariantCopyInd(GetValue, Default)
    End If
End Function

''
' Gets the HResult the current exception represents.
'
' @return The HResult.
' @remarks This is the same as Err.Number in most cases. Since this
' is a helper base class, it can hold any custom value needed for use
' by the Subclass.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Get HResult() As Long
    HResult = mHResult
End Property

''
' Sets the HResult the current exception represents.
'
' @param RHS The new HResult value.
' @remarks This is the same as Err.Number in most cases. Since this
' is a helper base class, it can hold any custom value needed for use
' by the Subclass.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Let HResult(ByVal RHS As Long)
    mHResult = RHS
End Property

''
' Gets the error message associated with the Subclass exception.
'
' @return A custom message set by the Subclass, or a default
' message of "An Error has occurred."
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Get Message() As String
    If cString.IsNull(mMessage) Then
        Message = Resources.GetString(Exception_WasThrown, mSubclassName)
    Else
        Message = mMessage
    End If
End Property

''
' Sets the error message associated with the Subclass exception.
'
' @param NewMessage The message to set the base to. Setting this to
' vbNullString will cause GetMessage to return the default message.
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Let Message(ByVal NewMessage As String)
    mMessage = NewMessage
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
    Set InnerException = mInnerException
End Property

''
' Sets the exception that caused the Subclassed exception to be thrown.
'
' @param RHS The exception that caused this Subclass exception to be thrown.
' @remarks when an exception is thrown and that exception causes another
' exception to be thrown, then the <b>InnerException</b> of the new <b>Exception</b>
' object should contain the exception that caused it to be thrown.
' <p>The Subclass exception should not allow the <b>InnerException</b> to be changed.</p>
'
Public Property Set InnerException(ByVal RHS As Exception)
    Set mInnerException = RHS
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
    Source = mSource
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
    mSource = RHS
End Property

''
' Gets a link to a help file associated with the exception.
'
' @return The Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Get HelpLink() As String
    HelpLink = mHelpLink
End Property

''
' Sets a link to a help file associated with the exception.
'
' @param RHS Set the Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Let HelpLink(ByVal RHS As String)
    mHelpLink = RHS
End Property

''
' Gets the original exception that caused the chain of exceptions to occur.
'
' @param Subclass The parent class that contains the <b>ExceptionBase</b> helper class.
' @return The <b>Exception</b> that caused the chain of exceptions to occur.
' @remarks If exceptions set their <b>InnerException</b> to a previously thrown
' exception, then a chain of exceptions can be created. Using this function will
' traverse that chain of exceptions until the original exception is reached. That
' exception with then be returned to the caller.
' <p>When an <b>InnerException</b> of Nothing is reached, then the Subclass is returned
' as the base exception because it did not have an inner exception, so it is assumed that
' the Subclass exception is the last in the chain and therefore the cause of the
' chain of exceptions being iterated.
'
Public Function GetBaseException() As Exception
    If mInnerException Is Nothing Then
        Set GetBaseException = Subclass
    Else
        Set GetBaseException = mInnerException.GetBaseException
    End If
End Function

''
' Returns the exception message prepended with the type name of the Subclass Exception.
'
' @param FormattedMessage This allows the Subclass to pass in a formatted version of the
' original message to be used in creating the resulting string. If the message is not
' passed in, then the internal <b>GetMessage</b> method will be used.
' @param AppTitle The name of the EXE, DLL or OCX that contains the exception Subclass.
' @param Subclass The parent Subclass object to derived the name from.
' @return A formatted message containing both the name of the parent subclass and
' the message itself.
' @remarks A general format might look like this:<br>
' VBCorLib.SystemException: An Error has occurred.
' <p>A listing of all inner exceptions will be included in the return value.</p>
'
Public Function ToString(Optional ByVal FormattedMessage As String)
    If Len(mSubclassName) > 0 Then ToString = mSubclassName & ": "
    
    If cString.IsNull(FormattedMessage) Then
        ToString = ToString & Message
    Else
        ToString = ToString & FormattedMessage
    End If
    
    If Not mInnerException Is Nothing Then ToString = ToString & vbCrLf & "---> " & mInnerException.ToString
End Function

''
' Provides a basic implementation of the Equals function of the <b>IObject</b> interface.
'
' @param Value The value to determine if is the same object instance as the Subclass exception.
' @param Subclass The Subclass exception object being compared to.
' @return Returns True if Value IS Subclass, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is Subclass)
    End If
End Function

''
' Provides a basic implementation of the GetHashcode function of the <b>IObject</b> interface.
'
' @param Subclass The parent Subclass object to derive the hashcode from.
' @return A 32-bit value used to help identify the Subclass object.
' @remarks The value generated is not unique across all hashcodes. Additional
' measures must be taken to find a unique value that happens to have the same
' hashcode as the Subclass object.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Subclass)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Subclass As Object, ByVal AppTitle As String, ByVal Message As String, ByVal InnerException As Exception, ByVal HResult As Long, ByVal Source As String, ByVal HelpLink As String)
    Call SetClassName(AppTitle, TypeName(Subclass))
    mSubclass = ObjPtr(CUnk(Subclass))
    mMessage = Message
    Set mInnerException = InnerException
    mHResult = HResult
    mSource = Source
    mHelpLink = HelpLink
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadDefaults()
    mHResult = DEF_HRESULT
End Sub

''
' This will attempt to build a string containing the application title
' and class name in a {AppTitle}.{Subclass} format.
'
' @param AppTitle The name of the component that contains the Subclass class.
' @param Subclass The exception class.
' @return A formatted string containing the applicationg title and class name.
' @remarks Possible formatted values are:
'
' {AppTitle}.{Subclass} - When AppTitle and Subclass are valid.
' {Subclass}            - When AppTitle is empty and Subclass is valid.
' {AppTitle}            - When Subclass is Nothing and AppTitle is valid.
' {Unknown}             - When AppTitle and Subclass are invalid.
'
Private Sub SetClassName(ByVal AppTitle As String, ByVal SubclassName As String)
    mSubclassName = AppTitle
    
    If Len(mSubclassName) > 0 Then
        mSubclassName = mSubclassName & "." & SubclassName
    Else
        mSubclassName = SubclassName
    End If
    
    If Len(mSubclassName) = 0 Then mSubclassName = "Unknown"
End Sub

''
' The mValues variable is only initialized upon first usage.
' We don't need to create a new Hashtable for exceptions that
' don't even store values.
'
' @return An initalized Hashtable used to store values.
'
Private Property Get Values() As Hashtable
    If mValues Is Nothing Then Set mValues = New Hashtable
    Set Values = mValues
End Property


Private Property Get Subclass() As Object
    Dim Unk As IUnknown
    ObjectPtr(Unk) = mSubclass
    Set Subclass = Unk
    ObjectPtr(Unk) = 0
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    Call LoadDefaults
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ' The PropertyBag does not persist vbNullString. It is converted to an empty string "".
        If Not CBool(.ReadProperty(PROP_ISNULLMESSAGE)) Then
            mMessage = .ReadProperty(PROP_MESSAGE)
        End If
        
        mHResult = .ReadProperty(PROP_HRESULT)
        mSource = .ReadProperty(PROP_SOURCE)
        mHelpLink = .ReadProperty(PROP_HELPLINK)
        Set mInnerException = .ReadProperty(PROP_INNEREXCEPTION, Nothing)   ' A default is actually required when an object might be Nothing.
        Set mValues = .ReadProperty(PROP_VALUES, Nothing)
        Set mData = .ReadProperty(PROP_DATA, Nothing)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_MESSAGE, mMessage)
        Call .WriteProperty(PROP_ISNULLMESSAGE, cString.IsNull(mMessage))
        Call .WriteProperty(PROP_HRESULT, mHResult)
        Call .WriteProperty(PROP_SOURCE, mSource)
        Call .WriteProperty(PROP_HELPLINK, mHelpLink)
        Call .WriteProperty(PROP_INNEREXCEPTION, mInnerException)
        
        ' Save the Hashtable only if it has been used.
        If Not mValues Is Nothing Then
            Call .WriteProperty(PROP_VALUES, mValues)
        End If
        
        If Not mData Is Nothing Then
            Call .WriteProperty(PROP_DATA, mData)
        End If
    End With
End Sub
