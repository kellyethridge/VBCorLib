VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hash Codes"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMD5 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   6015
   End
   Begin VB.TextBox txtRIPEMD160 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   4080
      Width           =   6015
   End
   Begin VB.TextBox txtSHA256 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtSHA384 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   6015
   End
   Begin VB.TextBox txtSHA512 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2640
      Width           =   6015
   End
   Begin VB.TextBox txtSHA1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label7 
      Caption         =   "Source:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "SHA-256:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "SHA-384:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "SHA-512:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "MD5:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "RIPEMD160:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "SHA-1:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSHA1 As New SHA1CryptoServiceProvider
Private mSHA256 As New SHA256Managed
Private mSHA384 As New SHA384Managed
Private mSHA512 As New SHA512Managed
Private mMD5 As New MD5CryptoServiceProvider
Private mRIPEMD160 As New RIPEMD160Managed



Private Sub txtSource_Change()
    Call ComputeHashes(txtSource.Text)
End Sub


Private Sub ComputeHashes(ByVal Text As String)
    Dim Bytes() As Byte
    Bytes = Encoding.UTF8.GetBytes(Text)
    
    Call DisplayHash(mSHA1.ComputeHash(Bytes), txtSHA1)
    Call DisplayHash(mSHA256.ComputeHash(Bytes), txtSHA256)
    Call DisplayHash(mSHA384.ComputeHash(Bytes), txtSHA384)
    Call DisplayHash(mSHA512.ComputeHash(Bytes), txtSHA512)
    Call DisplayHash(mMD5.ComputeHash(Bytes), txtMD5)
    Call DisplayHash(mRIPEMD160.ComputeHash(Bytes), txtRIPEMD160)
End Sub

Private Sub DisplayHash(ByRef Hash() As Byte, ByVal Box As TextBox)
    Dim sb As New StringBuilder
    sb.Length = 0
    
    Dim i As Long
    For i = 0 To UBound(Hash)
        If (i > 0) And ((i Mod 16) = 0) Then Call sb.AppendLine
        Call sb.AppendFormat("{0:X2}", Hash(i))
    Next i
    
    Box.Text = sb.ToString
End Sub
