VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResetCount 
      Caption         =   "Reset Count"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdStopTicker 
      Caption         =   "Stop Ticker"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartTicker 
      Caption         =   "Start Ticker"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "100"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Count:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Delay:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This version of the Ticker object uses typical
' Events to notify when time has elapsed.
'
Option Explicit

' Let's listen to the Ticker for an event.
Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private WithEvents mCounter As Counter
Attribute mCounter.VB_VarHelpID = -1

Private Sub cmdResetCount_Click()
    mCounter.Count = 0
End Sub

Private Sub cmdStartTicker_Click()
    ' We'll assume the user typed in an actual
    ' number and assign it as the new interval
    ' in milliseconds.
    
    ' When setting the Interval, if the Ticker is
    ' already running, it will be stopped and restarted.
    mTicker.Interval = Text1.Text
    
    ' If we weren't running to begin with, then setting
    ' the Interval won't start us back up, so check
    ' and start if we need to.
    If Not mTicker.Enabled Then mTicker.StartTicker
End Sub

Private Sub cmdStopTicker_Click()
    mTicker.StopTicker
End Sub

Private Sub Form_Load()
    Set mCounter = New Counter
    
    ' Use the constructor to easily create a new Ticker.
    Set mTicker = NewTicker(100, mCounter)
End Sub

Private Sub mCounter_Changed()
    Text2.Text = mCounter.Count
End Sub

' Respond to the Ticker event.
Private Sub mTicker_Elapsed(Data As Variant)
    ' By changing the value in mCount, the object
    ' will raise an event and the form will update.
    mCounter.Count = mCounter.Count + 1
End Sub

