VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' By Chris Lucas, cdl1051@earthlink.net, 20011204/20020607
' Thanks to Olaf for the class implementation concept

Option Explicit

Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)

Private Header1(5) As Long
Private Header2(5) As Long
Private SafeArray1() As Long
Private SafeArray2() As Long


Private Sub Class_Initialize()
    ' Set up our template for looking at strings
    Header1(0) = 1              ' Number of dimensions
    Header1(1) = 4              ' Bytes per element (long = 4)
    Header1(4) = &H7FFFFFFF     ' Array size
  
    ' Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray1), VarPtr(Header1(0)), 4
    
    ' Set up our template for look at search text
    Header2(0) = 1                 ' Number of dimensions
    Header2(1) = 4                 ' Bytes per element (long = 4)
    Header2(4) = &H7FFFFFFF        ' Array size
  
    ' Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray2), VarPtr(Header2(0)), 4
End Sub


Private Sub Class_Terminate()
    ' Make SafeArray1 once again use its own header
    ' If this code doesn't run the IDE will crash
    RtlMoveMemory ByVal ArrPtr(SafeArray1), 0&, 4
    RtlMoveMemory ByVal ArrPtr(SafeArray2), 0&, 4
End Sub


Friend Function EqualStrings(String1 As String, String2 As String, Optional Compare As VbCompareMethod = vbTextCompare) As Boolean
' By Chris Lucas, cdl1051@earthlink.net, 20011204
    Dim i&, SLen&, tmp1&, tmp2&, tmp3&, tmp4&, alt&
    
    SLen = LenB(String1)
    If SLen <> LenB(String2) Then Exit Function
    
    Header1(3) = StrPtr(String1): Header2(3) = StrPtr(String2)
    
    If Compare = vbTextCompare Then
        For i = 0 To SLen \ 4 - 1
            tmp1 = SafeArray1(i)
            tmp2 = (tmp1 And &HFFFF&)
            tmp3 = SafeArray2(i)
            tmp4 = (tmp3 And &HFFFF&)
            
            
            Select Case tmp2
                Case 97& To 122&: alt = tmp2 - 32
                Case 65& To 90&: alt = tmp2 + 32
                Case 49&: alt = 185
                Case 50&: alt = 178
                Case 51&: alt = 179
                Case 138&: alt = 154
                Case 140&: alt = 156
                Case 142&: alt = 158
                Case 154&: alt = 138
                Case 156&: alt = 140
                Case 158&: alt = 142
                Case 159&: alt = 255
                Case 178&: alt = 50
                Case 179&: alt = 51
                Case 185&: alt = 49
                Case 192& To 214&: alt = tmp2 + 32
                Case 216& To 222&: alt = tmp2 + 32
                Case 224& To 246&: alt = tmp2 - 32
                Case 248& To 254&: alt = tmp2 - 32
                Case 255&: alt = 376
                Case 338&: alt = 339
                Case 339&: alt = 338
                Case 352&: alt = 353
                Case 353&: alt = 352
                Case 376&: alt = 255
                Case 381&: alt = 382
                Case 382&: alt = 381
            End Select
            
            
            If alt <> tmp4 Then
                If tmp2 <> tmp4 Then Exit Function
            End If
            tmp2 = (tmp1 And &HFFFF0000)
            tmp4 = (tmp3 And &HFFFF0000)
            Select Case tmp2
                Case &H610000 To &H7A0000: alt = tmp2 - &H200000
                Case &H410000 To &H5A0000: alt = tmp2 + &H200000
                Case &H310000: alt = &HB90000
                Case &H320000: alt = &HB20000
                Case &H330000: alt = &HB30000
                Case &H8A0000: alt = &H9A0000
                Case &H8C0000: alt = &H9C0000
                Case &H8E0000: alt = &H9E0000
                Case &H9B0000: alt = &H8A0000
                Case &H9C0000: alt = &H8C0000
                Case &H9E0000: alt = &H8E0000
                Case &H9F0000: alt = &HFF0000
                Case &HB20000: alt = &H320000
                Case &HB30000: alt = &H970000
                Case &HB90000: alt = &H310000
                Case &HC00000 To &HD60000: alt = tmp2 + &H200000
                Case &HD80000 To &HDE0000: alt = tmp2 + &H200000
                Case &HE00000 To &HF60000: alt = tmp2 - &H200000
                Case &HF80000 To &HFE0000: alt = tmp2 - &H200000
                Case &HFF0000: alt = &H1780000
                Case &H1520000: alt = &H1530000
                Case &H1530000: alt = &H1520000
                Case &H1600000: alt = &H1610000
                Case &H1610000: alt = &H1600000
                Case &H1780000: alt = &HFF0000
                Case &H17D0000: alt = &H17E0000
                Case &H17E0000: alt = &H17D0000
            End Select
            If alt <> tmp4 Then
                If tmp2 <> tmp4 Then Exit Function
            End If
        Next i
        
        If (LenB(String1) \ 2 And 1) Then
            tmp2 = (SafeArray1(i) And &HFFFF&)
            tmp4 = (SafeArray2(i) And &HFFFF&)
            
            Select Case tmp2
                Case 97& To 122&: alt = tmp2 - 32
                Case 65& To 90&: alt = tmp2 + 32
                Case 49&: alt = 185
                Case 50&: alt = 178
                Case 51&: alt = 179
                Case 138&: alt = 154
                Case 140&: alt = 156
                Case 142&: alt = 158
                Case 154&: alt = 138
                Case 156&: alt = 140
                Case 158&: alt = 142
                Case 159&: alt = 255
                Case 178&: alt = 50
                Case 179&: alt = 51
                Case 185&: alt = 49
                Case 192& To 214&: alt = tmp2 + 32
                Case 216& To 222&: alt = tmp2 + 32
                Case 224& To 246&: alt = tmp2 - 32
                Case 248& To 254&: alt = tmp2 - 32
                Case 255&: alt = 376
                Case 338&: alt = 339
                Case 339&: alt = 338
                Case 352&: alt = 353
                Case 353&: alt = 352
                Case 376&: alt = 255
                Case 381&: alt = 382
                Case 382&: alt = 381
            End Select
            
            If tmp2 <> tmp4 Then
                If alt <> tmp4 Then Exit Function
            End If
        End If
    
        EqualStrings = True
    Else
        For i = 0 To SLen \ 4 - 1
            If SafeArray1(i) <> SafeArray2(i) Then Exit Function
        Next i

        If (LenB(String1) \ 2 And 1) Then
            If (SafeArray1(i) And &HFFFF&) <> (SafeArray2(i) And &HFFFF&) Then Exit Function
        End If

        EqualStrings = True
    End If
End Function


'Private Function EqualStringsBinaryCompare(String1 As String, String2 As String) As Boolean
'' By Chris Lucas, cdl1051@earthlink.net, 20020607
'
'    Dim i&, Len1&, Len2&, tmp&
'
'    ' Grab the string lengths
'    Len1 = LenB(String1) \ 2: Len2 = LenB(String2) \ 2
'
'    ' Make an informed decision as to whether we should continue
'    If Len1 <> Len2 Then GoTo BailOut
'
'    ' Compare the strings
'    Header1(3) = StrPtr(String1): Header2(3) = StrPtr(String2)
'    tmp = Len1 \ 2
'    ' The first two characters come cheap
'    If SafeArray1(i) <> SafeArray2(i) Then GoTo BailOut Else i = i + 1
'
'DoLoop:
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If SafeArray1(i) <> SafeArray2(i) Then GoTo NotEqual Else i = i + 1
'        If i <= tmp Then GoTo DoLoop
'
'NotEqual:
'    ' some characters don't match, but we need to check to
'    ' see if it happened after the end of the string, a
'    ' nasty side-effect of cascading ifs
'    If i > tmp Then EqualStringsBinaryCompare = True
'
'BailOut:
'    ' Lengths don't match, let's do absolutely nothing
'End Function
