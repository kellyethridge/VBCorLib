VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingTable"
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
'    Module: EncodingTable
'

Option Explicit

' The binary data read from the BinaryReader is
' in the same order as listed in this type. If more
' encoding items are to be added, they must be added
' in the same order using a BinaryWriter.
Private Type DataItem
    BodyName As String
    CodePage As Long
    EncodingName As String
    HeaderName As String
    IsBrowserDisplay As Boolean
    IsBrowserSave As Boolean
    IsMailNewsDisplay As Boolean
    IsMailNewsSave As Boolean
    WebName As String
    WindowsCodePage As Long
End Type

Private mEncodingData()     As DataItem
Private mIndexByCodePage    As Hashtable
Private mCodePageByName     As Hashtable


Friend Property Get BodyName(ByVal cp As Long) As String
    BodyName = mEncodingData(GetIndex(cp)).BodyName
End Property

Friend Property Get EncodingName(ByVal cp As Long) As String
    EncodingName = mEncodingData(GetIndex(cp)).EncodingName
End Property

Friend Property Get HeaderName(ByVal cp As Long) As String
    HeaderName = mEncodingData(GetIndex(cp)).HeaderName
End Property

Friend Property Get IsBrowserDisplay(ByVal cp As Long) As Boolean
    IsBrowserDisplay = mEncodingData(GetIndex(cp)).IsBrowserDisplay
End Property

Friend Property Get IsBrowserSave(ByVal cp As Long) As Boolean
    IsBrowserSave = mEncodingData(GetIndex(cp)).IsBrowserSave
End Property

Friend Property Get IsMailNewsDisplay(ByVal cp As Long) As Boolean
    IsMailNewsDisplay = mEncodingData(GetIndex(cp)).IsMailNewsDisplay
End Property

Friend Property Get IsMailNewsSave(ByVal cp As Long) As Boolean
    IsMailNewsSave = mEncodingData(GetIndex(cp)).IsMailNewsSave
End Property

Friend Property Get WebName(ByVal cp As Long) As String
    WebName = mEncodingData(GetIndex(cp)).WebName
End Property

Friend Property Get WindowsCodePage(ByVal cp As Long) As Long
    WindowsCodePage = mEncodingData(GetIndex(cp)).WindowsCodePage
End Property

Friend Function GetCodePage(ByVal Name As String) As Long
    Dim Ret As Variant
    Ret = mCodePageByName(LCase$(Name))
    If Not IsEmpty(Ret) Then
        GetCodePage = Ret
    Else
        Throw Cor.NewNotSupportedException("The CodePage is either invalid or not installed.")
    End If
End Function

Friend Function GetEncodings() As EncodingInfo()
    Dim Ret() As EncodingInfo
    ReDim Ret(0 To UBound(mEncodingData))
    
    Dim i As Long
    For i = 0 To UBound(Ret)
        Set Ret(i) = New EncodingInfo
        With mEncodingData(i)
            Call Ret(i).Init(.CodePage, .WebName, .EncodingName)
        End With
    Next i
    
    GetEncodings = Ret
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetIndex(ByVal cp As Long) As Long
    If mIndexByCodePage.Contains(cp) Then
        GetIndex = mIndexByCodePage(cp)
    Else
        Throw Cor.NewNotSupportedException("The CodePage is either invalid or not installed.")
    End If
End Function


' The byte stream is formatted using a BinaryWriter.
' The first item written is the number of encodings in
' the stream, as a vbLong. The order to read in each item
' of an encoding is the same as defined in DataItem at the top.
Private Sub LoadData()
    Dim b() As Byte
    b = LoadResData(101, "ENCODINGDATA")
    
    Dim ms As MemoryStream
    Set ms = Cor.NewMemoryStream(b, , , False)
    
    Dim Reader As BinaryReader
    Set Reader = Cor.NewBinaryReader(ms)
    
    Dim NumEncodings As Long
    NumEncodings = Reader.ReadLong
    ReDim mEncodingData(0 To NumEncodings - 1)
    
    Set mIndexByCodePage = New Hashtable
    Set mCodePageByName = New Hashtable
    
    Dim i As Long
    For i = 0 To NumEncodings - 1
        Call FillData(Reader, mEncodingData(i))
        Call mIndexByCodePage.Add(mEncodingData(i).CodePage, i)
        mCodePageByName(LCase$(mEncodingData(i).WebName)) = mEncodingData(i).CodePage
        mCodePageByName(LCase$(mEncodingData(i).BodyName)) = mEncodingData(i).CodePage
        mCodePageByName(LCase$(mEncodingData(i).EncodingName)) = mEncodingData(i).CodePage
    Next i

    Call Reader.CloseReader
    Call ms.CloseStream
End Sub

Private Sub FillData(ByVal Reader As BinaryReader, ByRef Data As DataItem)
    With Data
        .BodyName = Reader.ReadString
        .CodePage = Reader.ReadLong
        .EncodingName = Reader.ReadString
        .HeaderName = Reader.ReadString
        .IsBrowserDisplay = Reader.ReadBoolean
        .IsBrowserSave = Reader.ReadBoolean
        .IsMailNewsDisplay = Reader.ReadBoolean
        .IsMailNewsSave = Reader.ReadBoolean
        .WebName = Reader.ReadString
        .WindowsCodePage = Reader.ReadLong
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call LoadData
End Sub
