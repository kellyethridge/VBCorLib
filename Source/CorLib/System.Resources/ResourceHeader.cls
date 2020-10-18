VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'    Module: ResourceHeader
'

Option Explicit

Private mResType As ResourceTypes
Private mResTypeName As String
Private mResID As Long
Private mResName As String
Private mLangID As Long
Private mhModule As Long
Private mSize As Long
Private mOffset As Long


Friend Property Get ResType() As ResourceTypes
    ResType = mResType
End Property

Friend Property Get ResTypeName() As String
    ResTypeName = mResTypeName
End Property

Friend Property Get ResID() As Long
    ResID = mResID
End Property

Friend Property Get ResName() As String
    ResName = mResName
End Property

Friend Property Get LanguageID() As Long
    LanguageID = mLangID
End Property

Friend Property Get HModule() As Long
    HModule = mhModule
End Property

Friend Property Get Size() As Long
    Size = mSize
End Property

Friend Property Get Offset() As Long
    Offset = mOffset
End Property

Friend Sub Init(ByVal ResType As ResourceTypes, ByVal ResTypeName As String, ByVal ResID As Long, ByRef ResName As String, ByVal LangID As Long, ByVal HModule As Long)
    mResType = ResType
    mResTypeName = ResTypeName
    mResID = ResID
    mResName = ResName
    mLangID = LangID
    mhModule = HModule
    mSize = Size
    mOffset = Offset
End Sub
