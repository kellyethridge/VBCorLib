VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniResourceWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: IniResourceWriter
'

''
' Creates an entire section at once in an INI file.
'
' @remarks All values in the specified section are deleted, so there is no
' merging of the existing section key-value pairs and the new key-value pairs.
' Only the new key-value pairs will exist in the section once generated.
'
Option Explicit
Implements IObject
Implements IResourceWriter

Private mFileName   As String
Private mSection    As String
Private mValues     As Hashtable


''
' Returns the name of the INI file.
'
' @return The name of the INI file.
'
Public Property Get FileName() As String
    FileName = mFileName
End Property

''
' Returns the name of the section to write all the key-value pairs to.
'
' @return The name of the section in the INI file.
'
Public Property Get Section() As String
    Section = mSection
End Property

''
' Adds a key-value pair into the writer in preparation of creating an INI section.
'
' @param Key The identifier of the value to be written to the INI file section.
' @param Value The value to be written to the INI file section.
' @remarks If an object is passed in, then it must implement the <b>IObject</b>
' interface or an exception will be thrown. All other datatypes will be converted
' to their normal string value.
'
Public Sub AddResource(ByRef Key As String, ByRef Value As Variant)
    If Len(Key) = 0 Then _
        Error.Argument Argument_EmptyIniKey

    VerifyNotGenerated
    mValues.Add Key, Value
End Sub

''
' Creates a section within an INI file.
'
' @remarks All existing key-value pairs already in the specified section of the INI file will be removed.
' <p>The <b>Generate</b> method is called implicitly by the <b>CloseWriter</b> method if it is not called by your application code.</p>
' <p><b>Generate</b> can be called only once. After you call <b>Generate</b>, all <a href="IniResourceWriter.html">IniResourceWriter</a> methods other than <a href="IniResourceWriter.CloseWriter.html">CloseWriter</b> will throw an exception.</p>
Public Sub Generate()
    Dim sb      As New StringBuilder
    Dim Entry   As DictionaryEntry
    
    VerifyNotGenerated
    
    For Each Entry In mValues
       sb.Append Entry.Key
       sb.AppendChar vbEqualChar
       sb.Append Entry.Value
       sb.AppendString vbNullChar
    Next
    
    sb.AppendString vbNullChar
    
    If WritePrivateProfileSectionW(mSection, sb.ToString, mFileName) = BOOL_FALSE Then _
        IOError.WinIOError Err.LastDllError
    
    WritePrivateProfileStringW vbNullString, vbNullString, vbNullString, vbNullString
    Set mValues = Nothing
End Sub

''
' Saves the resources to the ini file and then closes it.
'
' @remarks <a href="IniResourceWriter.Generate.html">Generate</a> is called implicitly by <b>CloseWriter</b>, if required.
Public Sub CloseWriter()
    If Not mValues Is Nothing Then
        Generate
    End If
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef FileName As String, ByRef Section As String)
    If Len(FileName) = 0 Then _
        Error.Argument Argument_EmptyPath
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection
    
    mFileName = FileName
    mSection = Section
    Set mValues = New Hashtable
End Sub

Private Sub VerifyNotGenerated()
    If mValues Is Nothing Then
        Error.InvalidOperation InvalidOperation_ResourceWriterSaved
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceWriter
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceWriter_AddResource(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant)
    AddResource CStr(ResourceName), Convert.ToString(Value)
End Sub

Private Sub IResourceWriter_CloseWriter()
    CloseWriter
End Sub

Private Sub IResourceWriter_Generate()
    Generate
End Sub
