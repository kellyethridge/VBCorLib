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
' This version of the Ticker example shows how a
' callback method can be used to be notified of
' elapsed time instead of having to use an Event.
'
Option Explicit

' We aren't gonna listen for an Event from the Ticker.
Private mTicker As Ticker

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
    
    ' Create a new Ticker easily. We pass in the object we want
    ' the Ticker to be able to deal will without having to set
    ' some global variable that the Ticker can reach.
    '
    ' We also pass in the address of our callback function that
    ' will handle the elapsed time event.
    Set mTicker = NewTicker(100, mCounter, , AddressOf TickerCallback)
End Sub

Private Sub mCounter_Changed()
    Text2.Text = mCounter.Count
End Sub

' We have not mTicker_Elapsed event to handle.



