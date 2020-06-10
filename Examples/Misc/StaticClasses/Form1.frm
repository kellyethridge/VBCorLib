VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Some Available Functions"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "TimeZone"
      Height          =   3255
      Left            =   5400
      TabIndex        =   41
      Top             =   3960
      Width           =   4095
      Begin VB.TextBox txtDayLightEnd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtDayLightStart 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtTimeZoneOffset 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtStandardName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtDayLightName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtIsDayLightSavings 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label25 
         Caption         =   "DayLight End:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "DayLight Start:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "TimeZone Offset:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Standard Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "DayLight Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Is DayLight Savings:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MathExt"
      Height          =   3255
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   5175
      Begin VB.TextBox txtFloors 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox txtCeilings 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox txtDivRem 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtDegToRad 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtPI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtRadToDeg 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label19 
         Caption         =   "Floors:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Ceilings:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Divide With Remainder"
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Deg To Rad:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "PI:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Rad To Deg:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Environment"
      Height          =   3615
      Left            =   4320
      TabIndex        =   15
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtOperatingSystem 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox txtLogicalDrives 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtExpandVariable 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2160
         Width           =   5055
      End
      Begin VB.TextBox txtMemoryUsage 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtComputerName 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label13 
         Caption         =   "Operating System:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Logical Drives:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Expand Enivornment Variable"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Memory Usage: NT Only"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Computer Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Path"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtChangeExtension 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox txtExtension 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txtFileNameNoExt 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtRootDirectory 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtFullPath 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "c:\dir1\subdir2\filename.txt"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Change Extension:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Extension:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "FileName No Ext:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Root Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "FileName:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Full Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DemonstratePath()
    Dim FileName As String
    FileName = "FileName.txt"
    
    txtFullPath.Text = Path.GetFullPath(FileName)
    txtFileName.Text = Path.GetFileName(txtFullPath.Text)
    txtDirectory.Text = Path.GetDirectoryName(txtFullPath.Text)
    txtRootDirectory.Text = Path.GetPathRoot(txtFullPath.Text)
    txtFileNameNoExt.Text = Path.GetFileNameWithoutExtension(txtFullPath.Text)
    txtExtension.Text = Path.GetExtension(txtFullPath.Text)
    txtChangeExtension.Text = Path.ChangeExtension(txtFullPath.Text, "bin")
End Sub

Private Sub DemonstrateEnvironment()
    txtUserName.Text = Environment.UserName
    txtComputerName.Text = Environment.MachineName
    txtMemoryUsage.Text = CorString.Format("{0:n0}", Environment.WorkingSet)
    txtExpandVariable.Text = Environment.ExpandEnvironmentVariables("Path = %path%")
    txtLogicalDrives.Text = CorString.Join(", ", Environment.GetLogicalDrives)
    txtOperatingSystem.Text = Environment.OSVersion.ToString
End Sub

Private Sub DemonstrateMathExt()
    txtPI.Text = PI
    txtRadToDeg.Text = CorString.Format("{0} radians = {1} degrees", PI, CDeg(PI))
    txtDegToRad.Text = CorString.Format("{0} degrees = {1} radians", 90, CRad(90))
    
    Dim Quotient As Long
    Dim Remainder As Long
    Quotient = DivRem(13, 5, Remainder)
    txtDivRem.Text = CorString.Format("13/5 has a quotient of {0} and a remainder of {1}.", Quotient, Remainder)
    
    txtCeilings.Text = CorString.Format("Ceiling of {0} = {1} and {2} = {3}", 123.9, Ceiling(123.9), 123.1, Ceiling(123.1))
    txtFloors.Text = CorString.Format("Floor of {0} = {1} and {2} = {3}", 123.9, Floor(123.9), 123.1, Floor(123.1))
End Sub

Private Sub DemonstrateTimeZone()
    With TimeZone.CurrentTimeZone
        txtIsDayLightSavings.Text = .IsDayLightSavingTime(Now)
        txtDayLightName.Text = .DayLightName
        txtStandardName.Text = .StandardName
        txtTimeZoneOffset.Text = .GetDayLightChanges(Year(Now)).Delta.ToString
        txtDayLightStart.Text = .GetDayLightChanges(Year(Now)).StartTime.ToString
        txtDayLightEnd.Text = .GetDayLightChanges(Year(Now)).EndTime.ToString
    End With
End Sub

Private Sub Form_Load()
    DemonstratePath
    DemonstrateEnvironment
    DemonstrateMathExt
    DemonstrateTimeZone
End Sub
