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
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: ExceptionBase
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

Private Const PropMessage           As String = "Message"
Private Const PropIsNullMessage     As String = "IsNullMessage"
Private Const PropHResult           As String = "HResult"
Private Const PropErrorNumber       As String = "ErrorNumber"
Private Const PropSource            As String = "Source"
Private Const PropHelpLink          As String = "HelpLink"
Private Const PropInnerException    As String = "InnerException"
Private Const PropValues            As String = "Values"
Private Const PropData              As String = "Data"

Private mErrorNumber    As Long
Private mHResult        As Long
Private mMessage        As String
Private mSource         As String
Private mHelpLink       As String
Private mInnerException As Exception
Private mValues         As Hashtable
Private mSubclassPtr    As Long
Private mTypeName       As String
Private mData           As Hashtable


Public Property Get TypeName() As String
    TypeName = mTypeName
End Property

''
' Returns a key/value set of data that is associated to the exception.
'
' @return a key/value set of data that is associated to the exception.
' @remarks This is a public property that is used by the exception
' classes to allow user-defined information to be included with the exception.
'
Public Property Get Data() As IDictionary
    If mData Is Nothing Then
        Set mData = New Hashtable
    End If
    
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
Public Sub SetValue(ByRef Name As String, ByRef Value As Variant)
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
Public Function GetValue(ByRef Name As String, Optional ByRef Default As Variant) As Variant
    ' Steal the returned value from the Hashtable.
    Helper.MoveVariant GetValue, Values(Name)
    
    ' An empty value is the default return value for a key
    ' that does not exist in a Hashtable. It is also a valid
    ' value that may have been added to the Hashtable.
    ' So now we can check if the key exists. If it doesn't,
    ' then return the default value. This saves us from
    ' having to call into the Hashtable twice unless the
    ' returned value is Empty.
    If IsEmpty(GetValue) Then
        If Not Values.Contains(Name) Then
            VariantCopyInd GetValue, Default
        End If
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
' @param Value The new HResult value.
' @remarks This is the same as Err.Number in most cases. Since this
' is a helper base class, it can hold any custom value needed for use
' by the Subclass.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Let HResult(ByVal Value As Long)
    mHResult = Value
End Property

''
' Returns the current VB error number associated with this exception.
'
' @return The VB error number.
' @remarks This is equivelant to the Err.Number value.
'
Public Property Get ErrorNumber() As Long
    ErrorNumber = mErrorNumber
End Property

''
' Sets the VB error number for this exception instance.
'
' @param Value The new VB error number to set this instance to.
' @remarks This is equivelant to the Err.Number value.
'
Public Property Let ErrorNumber(ByVal Value As Long)
    mErrorNumber = Value
End Property

''
' Gets the error message associated with the Subclass exception.
'
' @return A custom message set by the Subclass.
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Get Message() As String
    Message = mMessage
End Property

''
' Sets the error message associated with the Subclass exception.
'
' @param Value The message to set the base to. Setting this to
' vbNullString will cause Environment.GetResourceString to return the default message.
' @remarks Generally this property is set to a meaningful message that
' is related to the exception that is being thrown. The message should
' be human readable.
' <p>This property can be set in the constructor <b>NewExceptionBase</b>.
'
Public Property Let Message(ByRef Value As String)
    mMessage = Value
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
' @param Value The exception that caused this Subclass exception to be thrown.
' @remarks when an exception is thrown and that exception causes another
' exception to be thrown, then the <b>InnerException</b> of the new <b>Exception</b>
' object should contain the exception that caused it to be thrown.
' <p>The Subclass exception should not allow the <b>InnerException</b> to be changed.</p>
'
Public Property Set InnerException(ByVal Value As Exception)
    Set mInnerException = Value
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
' @param Value A description of the source of the exception.
' @remarks The source of an exception generally will contain the name of
' the function being called when the exception was thrown. This is to help
' narrow down exactly where the exception had occurred.
'
Public Property Let Source(ByRef Value As String)
    mSource = Value
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
' @param Value Set the Uniform Resource Name (URN) or Uniform Resource Locator (URL).
' @remarks The return value, which represents a help file, is a URN or URL. For example, the HelpLink value could be:<br>
' "http://www.myhelpsite.com"
'
Public Property Let HelpLink(ByRef Value As String)
    mHelpLink = Value
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
' passed in, then the internal <b>Environment.GetResourceString</b> method will be used.
' @param AppTitle The name of the EXE, DLL or OCX that contains the exception Subclass.
' @param Subclass The parent Subclass object to derived the name from.
' @return A formatted message containing both the name of the parent subclass and
' the message itself.
' @remarks A general format might look like this:<br>
' VBCorLib.SystemException: An Error has occurred.
' <p>A listing of all inner exceptions will be included in the return value.</p>
'
Public Function ToString(Optional ByRef FormattedMessage As String)
    If Len(mTypeName) > 0 Then
        ToString = mTypeName & ": "
    End If
    
    If CorString.IsNull(FormattedMessage) Then
        ToString = ToString & Message
    Else
        ToString = ToString & FormattedMessage
    End If
    
    If Not mInnerException Is Nothing Then
        ToString = ToString & vbCrLf & "---> " & mInnerException.ToString
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Subclass As Exception, ByVal AppObject As Object)
    If Subclass Is Nothing Then _
        Error.ArgumentNull "Subclass"
    
    mTypeName = GetComponentName(AppObject) & VBA.TypeName(Subclass)
    mSubclassPtr = WeakPtr(Subclass)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetComponentName(ByVal AppObject As Object) As String
    On Error GoTo Catch
    
    If Not AppObject Is Nothing Then
        Dim Obj As App
        Set Obj = AppObject
        GetComponentName = Obj.Title & "."
    End If
    
Catch:
End Function

Private Property Get Values() As Hashtable
    If mValues Is Nothing Then
        Set mValues = New Hashtable
    End If
    
    Set Values = mValues
End Property

Private Property Get Subclass() As Exception
    Set Subclass = StrongPtr(mSubclassPtr)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ' The PropertyBag does not persist vbNullString. It is converted to an empty string "".
        If Not CBool(.ReadProperty(PropIsNullMessage)) Then
            mMessage = .ReadProperty(PropMessage)
        End If
        
        mHResult = .ReadProperty(PropHResult)
        mErrorNumber = .ReadProperty(PropErrorNumber)
        mSource = .ReadProperty(PropSource)
        mHelpLink = .ReadProperty(PropHelpLink)
        Set mInnerException = .ReadProperty(PropInnerException, Nothing)
        Set mValues = .ReadProperty(PropValues, Nothing)
        Set mData = .ReadProperty(PropData, Nothing)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropMessage, mMessage
        .WriteProperty PropIsNullMessage, CorString.IsNull(mMessage)
        .WriteProperty PropHResult, mHResult
        .WriteProperty PropErrorNumber, mErrorNumber
        .WriteProperty PropSource, mSource
        .WriteProperty PropHelpLink, mHelpLink
        .WriteProperty PropInnerException, mInnerException, Nothing
        .WriteProperty PropValues, mValues, Nothing
        .WriteProperty PropData, mData, Nothing
    End With
End Sub
