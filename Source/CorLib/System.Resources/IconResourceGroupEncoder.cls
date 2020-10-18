VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconResourceGroupEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2018 Kelly Ethridge
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
' Module: IconResourceGroupEncoder
'

''
' Encodes a <a href="PictureResourceGroup.html">PictureResourceGroup</a> object to a .RES byte array.
'
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder
Private Const RES_ICON          As Long = 1

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator


''
' Encodes a <a href="PictureResourceGroup.html">PictureResourceGroup</a> object.
'
' @param Value The <a href="PictureResourceGroup.html">PictureResourceGroup</a> to be encodes.
' @param ResourceName The identifier for the resource value. This can be a String or Number.
' @param ResourceType Identifies the type of resource this is. This value is ignored by this encoder.
' @param LanguageID The LocaleID the value is associated with.
' @return Returns <b>True</b> if the encoder was able to encode the value, <b>False</b> otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not a bitmap then the function returns False.
' <p>The <i>ResourceName</i> can be a String or a Numeric value.</p>
' <p>The LanguageID can be a <a href="CultureInfo.html">CultureInfo</a> object, a culture name, or an LCID. If the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    Dim Group   As PictureResourceGroup
    Dim Data    As New MemoryStream
    Dim Output  As BinaryWriter
    
    If Not IsPictureResourceGroup(Value) Then
        Exit Function
    End If
    
    Set Group = Value
    
    If Group.GroupType <> IconGroup Or Group.Count = 0 Then
        Exit Function
    End If
    
    Set Output = Cor.NewBinaryWriter(Data)
    Output.WriteValue CInt(0)
    Output.WriteValue CInt(RES_ICON)
    Output.WriteValue CInt(Group.Count)
    
    Dim i As Long
    For i = 0 To Group.Count - 1
        With Group(i)
            Output.WriteValue CByte(.Width)
            Output.WriteValue CByte(.Height)
            Output.WriteValue CInt(.Colors)
            Output.WriteValue CInt(.Planes)
            Output.WriteValue CInt(.BitCount)
            Output.WriteValue CLng(.Size)
            Output.WriteValue CInt(.ResourceId)
        End With
    Next
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceTypes.GroupIcon, GetLanguageID(LanguageID))
    
    mResources.Add Key, Data.ToArray
    Encode = True
End Function

''
' Retrieves the next encoded resource in the encoder.
'
' @param ReturnKey This is set to a <a href="ResourceKey.html">ResourceKey</a> object that uniquely identify the encoded resource.
' @param ReturnValue This will be set to a byte array.
' @return Returns <b>True</b> if an encoded resource was returned, <b>False</b> otherwise.
' @remarks This function can be called multiple times to retrieve multiple resources.
' As long as a resource is returned, <b>True</b> will be returned, otherwise <b>False</b> will be returned.
'
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue() As Byte) As Boolean
    If mIterator Is Nothing Then
        Set mIterator = mResources.GetEnumerator
    End If
    
    If mIterator.MoveNext Then
        Set ReturnKey = mIterator.Key
        ReturnValue = mIterator.Value
        GetEncodedResource = True
    End If
End Function

''
' Releases all currently encoded values.
'
Public Sub Reset()
    mResources.Clear
    Set mIterator = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceEncoder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceEncoder_Encode(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant) As Boolean
    IResourceEncoder_Encode = Encode(Value, ResourceName, ResourceType, LanguageID)
End Function

Private Function IResourceEncoder_GetEncodedResource(ReturnKey As ResourceKey, ReturnValue() As Byte) As Boolean
    IResourceEncoder_GetEncodedResource = GetEncodedResource(ReturnKey, ReturnValue)
End Function

Private Sub IResourceEncoder_Reset()
    Reset
End Sub
