VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SorterReferenceArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: SorterReferenceArray
'
Option Explicit

Private Const THRESHOLD As Long = 4

Private mItems() As Long
Private mpvData As Long
Private mComparer As ITwoArgReturnLong



Friend Sub Sort(ByVal pSA As Long, ByVal Left As Long, ByVal Right As Long)
    SAPtr(mItems) = pSA
    mpvData = MemLong(pSA + 12)
    On Error GoTo errTrap
    Select Case SafeArrayGetVartype(pSA)
        Case vbLong:                    Set mComparer = NewDelegator(AddressOf CompareLongs)
        Case vbString:                  Set mComparer = NewDelegator(AddressOf CompareStrings)
        Case vbObject, vbDataObject:    Set mComparer = NewDelegator(AddressOf CompareObjects)
    End Select
    QuickSort Left, Right
    InsertionSort Left, Right
errTrap:
    SAPtr(mItems) = 0
End Sub



Private Sub QuickSort(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    Dim pivot As Long
    Dim pPivot As Long

    If (Right - Left) <= THRESHOLD Then Exit Sub
    i = (Right + Left) / 2

    If mComparer.Call(mpvData + Left * 4, mpvData + i * 4) > 0 Then
        tmp = mItems(Left): mItems(Left) = mItems(i): mItems(i) = tmp
    End If
    If mComparer.Call(mpvData + Left * 4, mpvData + Right * 4) > 0 Then
        tmp = mItems(Left): mItems(Left) = mItems(Right): mItems(Right) = tmp
    End If
    If mComparer.Call(mpvData + i * 4, mpvData + Right * 4) > 0 Then
        tmp = mItems(i): mItems(i) = mItems(Right): mItems(Right) = tmp
    End If

    j = Right - 1
    tmp = mItems(i): mItems(i) = mItems(j): mItems(j) = tmp
    i = Left
    pivot = mItems(j)

    pPivot = VarPtr(pivot)
    Do
        Do: i = i + 1: Loop While mComparer.Call(mpvData + i * 4, pPivot) < 0
        Do: j = j - 1: Loop While mComparer.Call(mpvData + j * 4, pPivot) > 0

        If j < i Then Exit Do
        tmp = mItems(i): mItems(i) = mItems(j): mItems(j) = tmp
    Loop

    tmp = mItems(i): mItems(i) = mItems(Right - 1): mItems(Right - 1) = tmp

    QuickSort Left, j
    QuickSort i + 1, Right
End Sub

Private Sub InsertionSort(ByVal Left As Long, ByVal Right As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long
    Dim pTmp As Long
    
    pTmp = VarPtr(tmp)
    For i = Left + 1 To Right
        tmp = mItems(i)
        j = i
        
        Do While j > Left
            If mComparer.Call(mpvData + (j - 1) * 4, pTmp) < 0 Then Exit Do
            mItems(j) = mItems(j - 1)
            j = j - 1
        Loop
        
        mItems(j) = tmp
    Next i
End Sub

