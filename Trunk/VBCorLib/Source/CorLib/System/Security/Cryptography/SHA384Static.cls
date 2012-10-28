VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA384Static"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: SHA384Static
'

''
' Provides methods associated with the SHA-384 class.
'
Option Explicit

Private mStateH() As Long
Private mStateL() As Long
Private mStatePtr As DLong


''
' Creates a new instance of an SHA-384 implementation.
'
' @param AlgName The name of the implementation to create an instance of.
' @return A new SHA-384 compatible implementation.
' @remarks If not name is given, then a new SHA384Managed is returned.
'
Public Function Create(Optional ByVal AlgName As String = "SHA384") As SHA384
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get StatePtr() As DLong
    If mStatePtr.LoDWord = vbNullPtr Then
        mStateH = Cor.NewLongs(&HCBBB9D5D, &H629A292A, &H9159015A, &H152FECD8, &H67332667, &H8EB44A87, &HDB0C2E0D, &H47B5481D)
        mStateL = Cor.NewLongs(&HC1059ED8, &H367CD507, &H3070DD17, &HF70E5939, &HFFC00B31, &H68581511, &H64F98FA7, &HBEFA4FA4)
        
        mStatePtr.HiDWord = VarPtr(mStateH(0))
        mStatePtr.LoDWord = VarPtr(mStateL(0))
    End If
    
    StatePtr = mStatePtr
End Property
