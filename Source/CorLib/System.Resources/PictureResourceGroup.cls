VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PictureResourceGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: PictureResourceGroup
'

''
' This class represents a group of Icons or Cursors.
'
' @see PictureResourceGroupInfo
'
'@Folder("CorLib.System.Resources")
Option Explicit
Implements IObject


''
' The list of picture types supported.
'
' @param IconGroup Represents a group of icon pictures.
' @param CursorGroup Represents a group of cursor pictures.
'
Public Enum PictureGroupTypes
    IconGroup = ResourceTypes.IconResource
    CursorGroup = ResourceTypes.CursorResource
End Enum

Private mGroupType  As PictureGroupTypes
Private mItems      As New ArrayList


''
' Returns the number of pictures in the group.
'
' @return The number of pictures in the group.
'
Public Property Get Count() As Long
    Count = mItems.Count
End Property

''
' Returns the type of pictures in the group.
'
' @return The type of pictures in the group.
'
Public Property Get GroupType() As PictureGroupTypes
    GroupType = mGroupType
End Property

''
' Returns the info of a picture at the specific index.
'
' @param Index An index into the list of pictures in the group.
' @return A <a href="PictureResourceInfo.html">PictureResourceInfo</a> at the specified index.
'
Public Property Get Item(ByVal Index As Long) As PictureResourceInfo
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "40"
    Set Item = mItems(Index)
End Property

''
' Adds a picture to the group.
'
' @param Resource The picture to be added to the group.
' @remarks The picture resource value is not added to the resource. It is only
' used to obtain information about the picture. The actual picture
' resource needs to be added to the <a href="ResourceWriter.html">ResourceWriter</a> independently.
Public Sub Add(ByVal Resource As PictureResourceInfo)
    If Resource Is Nothing Then _
        Error.ArgumentNull "Resource"
    If Resource.ResourceType <> GroupType Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, IIf(GroupType = IconGroup, "Icon", "Cursor")), "Resource")
    
    mItems.Add Resource
End Sub

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compere equality with this instance.
' @return Returns True if the value equals this instance, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return Returns a hashcode value.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return A string representation of this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub InitFromBytes(ByRef Bytes() As Byte)
    Dim Source  As MemoryStream
    Dim Reader  As BinaryReader
    Dim Count   As Long
    
    Set Source = Cor.NewMemoryStream(Bytes, Writable:=False)
    Set Reader = Cor.NewBinaryReader(Source)
    Reader.ReadInt16 ' first two bytes are not used
    mGroupType = IIf(Reader.ReadInt16 = vbResIcon, PictureGroupTypes.IconGroup, PictureGroupTypes.CursorGroup)
    Count = Reader.ReadInt16
    
    Do While Count > 0
        Dim Info As PictureResourceInfo
        Set Info = New PictureResourceInfo
        Info.InitFromReader mGroupType, Reader
        mItems.Add Info
        Count = Count - 1
    Loop
End Sub

Friend Sub Init(ByVal GroupType As PictureGroupTypes)
    mGroupType = GroupType
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
