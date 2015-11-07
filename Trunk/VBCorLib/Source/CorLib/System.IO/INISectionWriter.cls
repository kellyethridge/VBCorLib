VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INISectionWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: INISectionWriter
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
Private mValues     As New Hashtable


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
Public Property Get SectionName() As String
    SectionName = mSection
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
Public Sub AddResource(ByVal Key As String, ByRef Value As Variant)
    mValues(Key) = Convert.ToString(Value)
End Sub

''
' Creates a section within an INI file.
'
' @remarks All existing key-value pairs already in the specified
' section of the INI file will be removed.
'
Public Sub Generate()
    Dim sb As New StringBuilder
    
    Dim de As DictionaryEntry
    For Each de In mValues
       Call sb.Append(de.Key)
       Call sb.AppendChar(corEqualChar)
       Call sb.Append(de.Value)
       Call sb.AppendChar(0)
    Next de
    Call sb.AppendChar(0)
    
    If Api.WritePrivateProfileSection(mSection, sb.ToString, mFileName) = BOOL_FALSE Then IOError.WinIOError Err.LastDllError
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal FileName As String, ByVal Section As String)
    If Len(FileName) = 0 Then _
        Throw Cor.NewArgumentException("FileName cannot be empty.", "FileName")
    If Len(Section) = 0 Then _
        Throw Cor.NewArgumentException("Section name cannot be empty.", "Section")
        
    mFileName = FileName
    mSection = Section
End Sub
    

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
'   IResourceWriter Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IResourceWriter_AddResource(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant)
    Call AddResource(Value, ResourceName)
End Sub

Private Sub IResourceWriter_CloseWriter()
' do nothing
End Sub

Private Sub IResourceWriter_Generate()
    Call Generate
End Sub
