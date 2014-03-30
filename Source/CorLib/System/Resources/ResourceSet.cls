VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceSet"
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
'    Module: ResourceSet
'

''
' Stores all the resources localized for one particular culture, ignoring all other cultures.
'
' @remarks This can be constructed using a .RES filename, a .RES formatted Stream, or
' an <b>IResourceReader</b> object.
' @see Constructors
' @see ResourceReader
' @see WinResourceReader
' @see ResourceManager
' @see Win32Resource
'
Option Explicit
Implements IObject
Implements IEnumerable

Private mResources  As Hashtable
Private mLanguageID As Long


''
' Returns the Culture ID for the specific set of resources.
'
' @return The culture ID.
'
Public Property Get LanguageID() As Long
    LanguageID = mLanguageID
End Property

''
' Returns the resource specified.
'
' @param ResourceName The resource identifier.
' @param ResourceType The type of resource to search for.
' @return The resource value, or Empty if the resource was not found.
' @remarks The ResType can be either a String of the type of resource or
' a ResourceTypes enum value.
'
Public Function GetObject(ByRef ResourceName As Variant, ByRef ResourceType As Variant) As Variant
    Call VerifyOpen
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceType, mLanguageID)
    
    Call Helper.MoveVariant(GetObject, mResources(Key))
End Function

''
' Returns a String resource from the set.
'
' @param ResourceName The string identifier.
' @return A resource string, or an empty string if not found.
'
Public Function GetString(ByRef ResourceName As Variant) As String
    Dim Ret As Variant
    Call Helper.MoveVariant(Ret, GetObject(ResourceName, ResourceTypes.StringResource))
    If IsEmpty(Ret) Then Exit Function
    GetString = Ret
End Function

''
' Returns an enumerator to iterate through all resources
' within the resource set.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
' The <i>Key</b> property returns the ID for the specific resource.
'
Public Function GetEnumerator() As IDictionaryEnumerator
    Call VerifyOpen
    Set GetEnumerator = mResources.GetEnumerator
End Function

''
' Returns an enumerator to iterate through all resources
' within the resource set.
'
' @return An enumerator.
' @remarks The enumerator returns values as <b>DictionaryEntry</b>
' objects. The value property in the <b>DictionaryEntry</b> object
' returns a <b>Win32Resource</b> object which contains details about
' the specific resource found in the .RES file.
' The <i>Key</b> property returns the ID for the specific resource.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Returns the name of the type of resource reader to use with
' this specifici <b>ResourceSet</b>
'
' @return The name of a resource reader type.
'
Public Function GetDefaultReader() As String
    GetDefaultReader = "VBCorLib.ResourceReader"
End Function

''
' Returns the name of the type of resource writer to use with
' this specifici <b>ResourceSet</b>
'
' @return The name of a resource writer type.
'
Public Function GetDefaultWriter() As String
    GetDefaultWriter = "VBCorLib.ResourceWriter"
End Function

''
' Closes the resource set.
'
Public Sub CloseResourceSet()
    Set mResources = Nothing
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef CultureID As Variant, ByRef Source As Variant)
    mLanguageID = GetLanguageID(CultureID)
    Call LoadResources(GetReader(Source))
End Sub

Private Sub LoadResources(ByVal Reader As IResourceReader)
    Set mResources = New Hashtable
    
    Dim de As DictionaryEntry
    For Each de In Reader
        Dim Key As ResourceKey
        Set Key = de.Key
        If Key.LanguageID = mLanguageID Then
            Call mResources.Add(de.Key, de.Value)
        End If
    Next de
End Sub

Friend Sub AddResource(ByVal Key As ResourceKey, ByRef Data As Variant)
    If mResources Is Nothing Then Set mResources = New Hashtable
    Call mResources.Add(Key, Data)
End Sub

Friend Function GetResourceSet() As Hashtable
    Set GetResourceSet = mResources
End Function

Friend Property Let LanguageID(ByVal RHS As Long)
    mLanguageID = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyOpen()
    If mResources Is Nothing Then Throw Cor.NewInvalidOperationException("ResourceSet is closed.")
End Sub

Private Function GetReader(ByRef Source As Variant) As IResourceReader
    If IsObject(Source) Then
        If Source Is Nothing Then _
            Throw Cor.NewArgumentNullException("Source")
            
        If TypeOf Source Is IResourceReader Then
            Set GetReader = Source
        End If
    End If
    If GetReader Is Nothing Then Set GetReader = Cor.NewResourceReader(Source)
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
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function
