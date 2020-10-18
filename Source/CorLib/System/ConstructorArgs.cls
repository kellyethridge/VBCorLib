VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstructorArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: ConstructorArgs
'

Option Explicit

Private mArgs() As Variant



Friend Sub Init(ParamArray Args() As Variant)
    mArgs = Args
End Sub

Friend Property Get IsEmpty() As Boolean
    IsEmpty = (Count = 0)
End Property

Friend Property Get Count() As Long
    If cArray.IsNull(mArgs) Then Exit Property
    
    Count = UBound(mArgs) + 1
End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Call VariantCopy(Item, mArgs(Index))
End Property

Friend Property Get ArgType(ByVal Index As Long) As VbVarType
    ArgType = VarType(mArgs(Index))
End Property

Friend Sub Clear()
    Erase mArgs
End Sub
