VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResCursorEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: ResCurcorEncoder
'

''
' Encodes a StdPicture Cursor to byte data used in .RES files.
'
' @see ResourceWriter
'
Option Explicit
Implements IResourceEncoder

Private Const SIZEOF_ICONDIR_AND_1_ICONDIRENTRY As Long = 22
Private Const SIZEOF_HOTSPOT                    As Long = 4

Private mResources  As New Hashtable
Private mIterator   As IDictionaryEnumerator


''
' Encodes the StdPicture cursor.
'
' @param Value The StdPicture to be encoded.
' @param ResourceName The identifier for the resource value. This can be a String or Number.
' @param ResourceType Identifies the type of resource this is. If this is not supplied, then
' the type is derived from the value.
' @param LanguageID The LocaleID the value is associated with.
' @return Returns True if the encoder was able to encode the value, False otherwise.
' @remarks Any value can be passed in. No exception is thrown. If the value is not
' a bitmap then the function returns False.
' <p>The <b>ResourceName</b> can be a String or a Numeric value.</p>
' <p>The LanguageID can be a <b>CultureInfo</b> object, a culture name, or an LCID. If
' the LanguageID is not supplied, then the current culture is used.</p>
'
Public Function Encode(ByRef Value As Variant, ByRef ResourceName As Variant, Optional ByRef ResourceType As Variant, Optional ByRef LanguageID As Variant) As Boolean
    If IsMissing(ResourceType) Then
        ' If the resource type is missing, then attempt
        ' to determine the type from the value.
        If Not IsObject(Value) Then Exit Function
        If Value Is Nothing Then Exit Function
        If Not TypeOf Value Is IPicture Then Exit Function
    Else
        If Not IsInteger(ResourceType) Then Exit Function
        If ResourceType <> ResourceTypes.CursorResource Then Exit Function
    End If
    
    Dim Pic As IPicture
    Set Pic = Value
    
    ' This is true for both Icons and Cursors.
    If Pic.Type <> PICTYPE_ICON Then Exit Function
    
    Dim Info As ICONINFO
    Call GetIconInfo(Pic.Handle, Info)
    
    ' This will tell us if the IPicture is a cursor or an icon.
    If Info.fIcon = BOOL_TRUE Then Exit Function
    
    ' Prepare to have the picture save itself out.
    Dim Stream As IStream
    Call CreateStreamOnHGlobal(0, True, Stream)
    
    ' Save the picture, producing the bytes we need.
    Dim BytesWritten As Long
    Call Pic.SaveAsFile(ByVal ObjPtr(Stream), True, BytesWritten)
    
    ' We don't include the size of the BITMAPFILEHEADER.
    BytesWritten = BytesWritten - SIZEOF_ICONDIR_AND_1_ICONDIRENTRY
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To SIZEOF_HOTSPOT + BytesWritten - 1)
    
    ' Move to the first byte past the SIZEOF_ICONDIR_AND_1_ICONDIRENTRY structure.
    ' 0.0022 as vbCurrency equals 22 as 64bit.
    Call Stream.Seek(0.0001@ * SIZEOF_ICONDIR_AND_1_ICONDIRENTRY, STREAM_SEEK_SET, 0@)
    
    ' And read the bytes back out of the stream.
    Call Stream.Read(VarPtr(Bytes(SIZEOF_HOTSPOT)), BytesWritten, 0)
    
    ' Now assign the hot spot values obtained earlier.
    ' Even though the API defines hotspots as 32bit values,
    ' the .RES file defines them at 16bit values.
    AsWord(Bytes(0)) = Info.xHotSpot
    AsWord(Bytes(2)) = Info.yHotSpot
    
    Dim Key As ResourceKey
    Set Key = Cor.NewResourceKey(ResourceName, ResourceTypes.CursorResource, GetLanguageID(LanguageID))
    
    Call mResources.Add(Key, Bytes)
    
    Encode = True
End Function

''
' Retrieves the next encoded resource in the encoder.
'
' @param ReturnKey This is set to a <b>ResourceKey</b> object that
' uniquely identify the encoded resource.
' @param ReturnValue This will be set to a byte array.
' @return Returns True if an encoded resource was returned, False otherwise.
' @remarks This function can be called multiple times to retrieve multiple resources.
' As long as a resource is returned, True will be returned, otherwise False will be returned.
'
Public Function GetEncodedResource(ByRef ReturnKey As ResourceKey, ByRef ReturnValue As Variant) As Boolean
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
    Call mResources.Clear
    Set mIterator = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IResourceEncoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceEncoder_Encode(Value As Variant, ResourceName As Variant, Optional ResourceType As Variant, Optional LanguageID As Variant) As Boolean
    IResourceEncoder_Encode = Encode(Value, ResourceName, ResourceType, LanguageID)
End Function

Private Function IResourceEncoder_GetEncodedResource(ReturnKey As ResourceKey, ReturnValue As Variant) As Boolean
    IResourceEncoder_GetEncodedResource = GetEncodedResource(ReturnKey, ReturnValue)
End Function

Private Sub IResourceEncoder_Reset()
    Call Reset
End Sub


