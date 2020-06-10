VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Some Formatted Numbers"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Caption         =   "Hex (X)"
      Height          =   2055
      Left            =   5760
      TabIndex        =   44
      Top             =   960
      Width           =   2655
      Begin VB.TextBox txtHex8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   52
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtHex2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   50
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtHexLower 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtHex 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "{0:X8} :"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "{0:X2} :"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "{0:x} :"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "{0:X} :"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Float"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   43
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "64-Integer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   42
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   "Number (N)"
      Height          =   1815
      Left            =   3120
      TabIndex        =   35
      Top             =   4560
      Width           =   5295
      Begin VB.TextBox txtNumber4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox txtNumber0 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtNumber 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label18 
         Caption         =   "{0:N4} :"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "{0:N0} :"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "{0:N} :"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Currency (C)"
      Height          =   1335
      Left            =   5760
      TabIndex        =   30
      Top             =   3120
      Width           =   2655
      Begin VB.TextBox txtCurrency3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtCurrency 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "{0:C3} :"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "{0:C} :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fixed (F)"
      Height          =   1935
      Left            =   3120
      TabIndex        =   23
      Top             =   2520
      Width           =   2415
      Begin VB.TextBox txtFixed0 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtFixed4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFixed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "{0:F0} :"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "{0:F4} :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "{0:F} :"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scientific (E - 123456)"
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   2775
      Begin VB.TextBox txtExponent4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtExponentLower 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtExponent 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "{0:E4}:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "{0:e}:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "{0:E}:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Decimal (D - 123456)"
      Height          =   1455
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Width           =   2415
      Begin VB.TextBox txtDecimal10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtDecimal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "{0:D10} :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "{0:D} :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General (G)"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2775
      Begin VB.TextBox txtGeneral3Lower 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtGeneral3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtGeneral10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtGeneral 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "{0:g3} :"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "{0:G3} :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "{0:G10} :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "{0:G} :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   12360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label11 
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mNumber As Variant

Private Sub UpdateGeneral()
    ' When formating a number, "G"eneral is the default
    ' formatting method used. The default formatting for
    ' General is to show the entire number.
    txtGeneral.Text = CorString.Format("{0:G}", mNumber) ' or {0:G0} or {0}
    
    ' Specifying 10 after the "G" indicates that we want
    ' to see 10 digits maximum. If more than 10 digits occur
    ' then the number is formatted in scientific notation.
    txtGeneral10.Text = CorString.Format("{0:G10}", mNumber)
    
    ' Specifying 3 after the "G" indicates that we want
    ' to see 3 significant digits. If more than 3 digits occur
    ' then the number is formatted in scientific notation.
    txtGeneral3.Text = CorString.Format("{0:G3}", mNumber)
    
    ' This also specifies 3 and works identical to the above
    ' formatting. However the "G" is in lowercase. This indicates
    ' that the "E" in the scientific notation is to be lowercase.
    txtGeneral3Lower.Text = CorString.Format("{0:g3}", mNumber)
End Sub

Private Sub UpdateDecimal()
    ' Decimal is only support on integer numbers only.
    If Option1.Value = True Then
        ' Specifying a "D" alone indicates to just show the number.
        txtDecimal.Text = CorString.Format("{0:D}", mNumber)
        
        ' Specifying a 10 after the "D" indicates we want a total
        ' of 10 digits displayed. If there are not enough digits
        ' in the number, then pad the beginning with zeros to
        ' fill in the missing digits.
        txtDecimal10.Text = CorString.Format("{0:D10}", mNumber)
    Else
        txtDecimal.Text = "N/A"
        txtDecimal10.Text = "N/A"
    End If
End Sub

Private Sub UpdateScientific()
    ' Specifies to display the number in default scientific notation.
    ' The default it to show 7 total digits (0.000000E+000). The "E"
    ' is also in uppercase by default.
    txtExponent.Text = CorString.Format("{0:E}", mNumber)
    
    ' This works identical to the above format except that the "E"
    ' is to be displayed in lowercase.
    txtExponentLower.Text = CorString.Format("{0:e}", mNumber)
    
    ' The 4 after the "E" indicates that we want to see exactly
    ' 4 digits after the decimal place (0.0000E+000).
    txtExponent4.Text = CorString.Format("{0:E4}", mNumber)
End Sub

Private Sub UpdateFixed()
    ' A Fixed number has a fixed set of digits following
    ' the decimal point. If the number is an integer, then
    ' zeros are placed after the decimal point. The default
    ' number of digits to display after the decimal point is 2.
    txtFixed.Text = CorString.Format("{0:F}", mNumber)
    
    ' Specifies that 4 digits are to be displayed after the decimal point.
    ' If there are not enough digits, then zeros fill in the missing digits.
    txtFixed4.Text = CorString.Format("{0:F4}", mNumber)
    
    ' Specifying 0 after then "F" indicates that there are to
    ' be no digits displayed after the decimal point. The
    ' decimal point is also dropped.
    txtFixed0.Text = CorString.Format("{0:F0}", mNumber)
End Sub

Private Sub UpdateCurrency()
    ' Specifying "C" indicating the number should be formatted using the
    ' current Currency formatting valued. This can change depending
    ' on regional settings and user overrides.
    txtCurrency.Text = CorString.Format("{0:C}", mNumber)
    
    ' Specifying 3 after the "C" indicates that the value should
    ' contain exactly 3 digits after the decimal point. All other
    ' formatting is dependant on region and user settings.
    txtCurrency3.Text = CorString.Format("{0:C3}", mNumber)
End Sub

Private Sub UpdateNumber()
    ' Specifying "N" indicates that the number should be formatting
    ' using a decimal point if necessary and grouping the digits
    ' together based on the group settings. Usually this is set
    ' to grouping each set of 3 digits, separating the groups with a comma.
    ' A default of 2 digits will follow the decimal point.
    txtNumber.Text = CorString.Format("{0:N}", mNumber)
    
    ' The 0 following the "N" indicates that no digits should follow
    ' the decimal point. The decimal point will also be dropped.
    txtNumber0.Text = CorString.Format("{0:N0}", mNumber)
    
    ' The 4 following the "N" indicates that exactly 4 digits will
    ' follow the decimal point. Zeros will be used to fill in for
    ' missing digits.
    txtNumber4.Text = CorString.Format("{0:N4}", mNumber)
End Sub

Private Sub UpdateHex()
    ' Hex formatting is only supported for integer values.
    If Option1.Value = True Then
        ' Formats the value in hexidecimal notation. If the
        ' value is a vbInteger, then it will be a max of 4 characters.
        txtHex.Text = CorString.Format("{0:X}", mNumber)
        
        ' This works identical to the method above, however the
        ' alpha digits are output in lowercase.
        txtHexLower.Text = CorString.Format("{0:x}", mNumber)
        
        ' The 2 after the "X" indicates that the output should
        ' contain atleast 2 characters. If there are not enough
        ' digits, then the number is preceeded with zeros to
        ' fill in for the missing digits.
        txtHex2.Text = CorString.Format("{0:X2}", mNumber)
        
        ' This works identical to the method above. In this case
        ' the minimum digits displayed is 8. If there are not enough
        ' then zeros are used to fill in.
        txtHex8.Text = CorString.Format("{0:X8}", mNumber)
    Else
        txtHex.Text = "N/A"
        txtHexLower.Text = "N/A"
        txtHex2.Text = "N/A"
        txtHex8.Text = "N/A"
    End If
End Sub

Private Sub Option1_Click()
    Update
End Sub

Private Sub Option2_Click()
    Update
End Sub

Private Sub txtInput_Change()
    Update
End Sub

Private Sub Update()
    On Error Resume Next
    If Option1.Value Then
        mNumber = CInt64(txtInput.Text)
    Else
        mNumber = CDbl(Val(txtInput.Text))
    End If
    
    UpdateGeneral
    UpdateDecimal
    UpdateScientific
    UpdateFixed
    UpdateCurrency
    UpdateNumber
    UpdateHex
End Sub
