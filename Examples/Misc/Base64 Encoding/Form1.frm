VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowseNewFile 
      Caption         =   "Browse"
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   12
      Top             =   9360
      Width           =   6495
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode To New File"
      Height          =   495
      Left            =   8880
      TabIndex        =   11
      Top             =   9240
      Width           =   2175
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode Base64"
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   8640
      Width           =   2175
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7815
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13785
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   9840
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6000
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "New File:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label lblEncodedLength 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Encoded Length:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label lblFileLength 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "File Length:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "File:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example will encode a selected file into a Base-64 encoding
' using the MIME 76 character linebreak rule.
'
' This also allows for the characters to be decoded back into a file.
'
Option Explicit

' Browses for a file to be encoded.
Private Sub cmdBrowse_Click()
    On Error GoTo errTrap
    With CD
        .CancelError = True
        .ShowOpen
        Text1.Text = .FileName
    End With
errTrap:
End Sub

' Browses for a new file to ouput the decode string to.
Private Sub cmdBrowseNewFile_Click()
    On Error GoTo errTrap
    With CD
        .CancelError = True
        .ShowOpen
        Text2.Text = .FileName
    End With
errTrap:
End Sub

' This will decode a Base64 string into a file.
Private Sub cmdDecode_Click()
    ' Do we want to overwrite an existing file?
    If File.Exists(Text2.Text) Then
        If MsgBox("File already exists, Overwrite?", vbYesNo) = vbNo Then Exit Sub
        File.Delete Text2.Text
    End If
    
    ' Get the string out of the richtextbox, since it
    ' seems to be a pig when it comes to large strings.
    ' (No offense to pigs)
    Dim s As String
    s = RichTextBox1.Text
    
    ' Setup our timer to see how long it takes.
    Dim sw As StopWatch
    Set sw = StopWatch.StartNew
    
    ' Decode the Base-64 string back into the binary data.
    Dim b() As Byte
    b = Convert.FromBase64String(s)
    
    ' And show the time elapsed.
    lblTime.Caption = sw.ElapsedMilliseconds
    
    File.WriteAllbytes Text2.Text, b
End Sub

' Encodes a selected file into a Base64 string.
Private Sub cmdEncode_Click()
    ' Lets check if the file exists and let the
    ' user know if we couldn't find it.
    If Not File.Exists(Text1.Text) Then
        MsgBox "File does not exist.", vbExclamation + vbOKOnly, "File Not Found."
        Exit Sub
    End If

    ' Create our timing object to tell us how long it took to encode.
    Dim sw As StopWatch
    Set sw = StopWatch.StartNew
    
    Dim b() As Byte
    Dim s As String
    
    b = File.ReadAllBytes(Text1.Text)
    s = Convert.ToBase64String(b)
    
    ' let the world know how long the encoding took.
    lblTime.Caption = sw.ElapsedMilliseconds
    
    ' And show other stats about the file and encoded string.
    lblFileLength = FileLen(Text1.Text)
    lblEncodedLength.Caption = Len(s)
    
    ' Finally show the encoded string in the textbox.
    RichTextBox1.Text = s
End Sub

