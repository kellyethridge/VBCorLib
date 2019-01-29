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
' Module: ResourceSet
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
    VerifyOpen
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceType, mLanguageID)
    
    Helper.MoveVariant GetObject, mResources(Key)
End Function

''
' Returns a String resource from the set.
'
' @param ResourceName The string identifier.
' @return A resource string, or an empty string if not found.
'
Public Function GetString(ByRef ResourceName As Variant) As String
    Dim Ret As Variant
    Helper.MoveVariant Ret, GetObject(ResourceName, ResourceTypes.StringResource)
    
    If IsEmpty(Ret) Then
        Exit Function
    End If
    
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
    VerifyOpen
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
    GetDefaultReader = "CorLib.ResourceReader"
End Function

''
' Returns the name of the type of resource writer to use with
' this specifici <b>ResourceSet</b>
'
' @return The name of a resource writer type.
'
Public Function GetDefaultWriter() As String
    GetDefaultWriter = "CorLib.ResourceWriter"
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
Friend Sub Init(ByRef CultureID As Variant, ByRef Source As Variant)
    mLanguageID = GetLanguageID(CultureID)
    LoadResources GetReader(Source)
End Sub

Private Sub LoadResources(ByVal Reader As IResourceReader)
    Set mResources = New Hashtable
    
    Dim de As DictionaryEntry
    For Each de In Reader
        Dim Key As ResourceKey
        Set Key = de.Key
        
        If Key.LanguageID = mLanguageID Then
            mResources.Add de.Key, de.Value
        End If
    Next de
End Sub

Friend Sub AddResource(ByVal Key As ResourceKey, ByRef Data As Variant)
    If mResources Is Nothing Then
        Set mResources = New Hashtable
    End If
    
    mResources.Add Key, Data
End Sub

Friend Function GetResourceSet() As Hashtable
    Set GetResourceSet = mResources
End Function

Friend Property Let LanguageID(ByVal Value As Long)
    mLanguageID = Value
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyOpen()
    If mResources Is Nothing Then
        Throw Cor.NewInvalidOperationException("ResourceSet is closed.")
    End If
End Sub

Private Function GetReader(ByRef Source As Variant) As IResourceReader
    If IsObject(Source) Then
        If Source Is Nothing Then _
            Error.ArgumentNull "Source"
            
        If TypeOf Source Is IResourceReader Then
            Set GetReader = Source
        End If
    End If
    
    If GetReader Is Nothing Then
        Set GetReader = Cor.NewResourceReader(Source)
    End If
End Function


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
'   IEnumerable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function
