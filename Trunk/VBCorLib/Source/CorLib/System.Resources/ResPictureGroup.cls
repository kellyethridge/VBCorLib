VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResPictureGroup"
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
'    Module: ResPictureGroup
'

''
' This class represents a group of Icons or Cursors.
'
' @see ResPictureGroupInfo
'
Option Explicit
Implements IObject

Private Const SIZEOF_NEWHEADER  As Long = 6
Private Const SIZEOF_RESDIR     As Long = 14
Private Const RES_ICON          As Long = 1

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
' @return A <b>ResPictureGroupInfo</b> at the specified index.
'
Public Property Get Item(ByVal Index As Long) As ResPictureGroupInfo
Attribute Item.VB_UserMemId = 0
    Set Item = mItems(Index)
End Property

''
' Adds a picture to the group.
'
' @param ResourceID The numeric ID of the picture resource to be included in the group.
' @param Pic The picture to be added to the group.
' @remarks The Pic value is not added to the resource. It is only
' used to obtain information about the picture. The actual picture
' resource needs to be added to the <b>ResourceWriter</b> independently.
'
Public Sub Add(ByVal ResourceId As Long, ByVal Pic As IPicture)
    If Pic Is Nothing Then _
        Throw New ArgumentNullException
    If Pic.Type <> PICTYPE_ICON Then _
        Throw Cor.NewArgumentException("Only Icons and Cursors are supported.", "Pic")
    
    Dim Info As New ResPictureGroupInfo
    Call Info.Init(ResourceId, Pic, mGroupType)
    Call mItems.Add(Info)
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub InitFromBytes(ByRef Bytes() As Byte)
    mGroupType = IIf(AsWord(Bytes(2)) = RES_ICON, PictureGroupTypes.IconGroup, PictureGroupTypes.CursorGroup)
    
    Dim Count As Long
    Count = AsWord(Bytes(4))
    
    Dim Index As Long
    Index = SIZEOF_NEWHEADER   ' Skip the header
    
    Dim i As Long
    For i = 1 To Count
        Dim Info As ResPictureGroupInfo
        Set Info = New ResPictureGroupInfo
        
        Index = Info.Parse(mGroupType, Index, Bytes)
        Call mItems.Add(Info)
    Next i
End Sub

Friend Sub InitNew(ByVal GroupType As PictureGroupTypes)
    mGroupType = GroupType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
