Attribute VB_Name = "modMain"
' This example simply moves a block of characters around
' the console using the various console methods. It makes
' a block of characters appear to be bouncing around.
'
' *** WARNING ***
' Remember to never click on the Console close button, or
' press the End button in the IDE, or the application will crash.
'
Option Explicit
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Main()
    ' Create the original block of characters in the upper-left corner.
    Console.FillBufferArea 0, 0, 5, 5, "*", Cyan
    
    ' Get rid of the cursor.
    Console.CursorVisible = False
    
    ' Set our default delay and display the console title.
    Dim Delay As Long
    Delay = 50
    SetTitle Delay
    
    ' Set our initial block movement directions.
    Dim dx As Long
    Dim dy As Long
    dx = 1
    dy = 1

    Dim x As Long
    Dim y As Long
    Do
        ' Calculate the next coordinates for
        ' our block of characters.
        Dim NewX As Long
        Dim NewY As Long
        NewX = x + dx
        NewY = y + dy

        ' Move the block of characters from the original
        ' location to the new calculated location.
        Console.MoveBufferArea x, y, 5, 5, NewX, NewY

        ' Set the x to the new calculated location and check
        ' to see if we hit the left or right side of the
        ' console. If we did, then reverse the direction.
        x = NewX
        If x = 0 Or x = 75 Then dx = 0 - dx

        ' Set the y to the new calculated location and check
        ' to see if we hit the top or bottom  of the
        ' console. If we did, then reverse the direction.
        y = NewY
        If y = 0 Or y = 20 Then dy = 0 - dy

        ' Check to see if a key has been pressed in the console.
        ' This method does not block, it only notifies.
        If Console.KeyAvailable Then
            ' We have a key ready for retrieval.

            ' Retrieve the key from the console. We want to Intercept
            ' the key to prevent it from being displayed in the console.
            '
            ' Perform a Select statement on the key code.
            Select Case Console.ReadKey(True).Key

                ' Pressed the Escape key, so you must want out.
                Case ConsoleKey.EscapeKey
                    Exit Do

                ' Cursor down will lower the delay, speeding
                ' up the block movment. We don't allow a delay
                ' of less than 10ms. Change the title to reflect
                ' the new delay selected.
                Case ConsoleKey.DownArrowKey
                    Delay = Delay - 1
                    If Delay < 10 Then Delay = 10
                    SetTitle Delay
                    
                ' Cursor up will raise the delay, slowing
                ' down the block movment. We don't allow a delay
                ' of more than 200ms. Change the title to reflect
                ' the new delay selected.
                Case ConsoleKey.UpArrowKey
                    Delay = Delay + 1
                    If Delay > 200 Then Delay = 200
                    SetTitle Delay
            End Select
        End If
        Sleep Delay
    Loop
End Sub

' Set the title of the console.
Private Sub SetTitle(ByVal Delay As Long)
    Console.Title = CorString.Format("Bouncing Square - (Delay of {0}) - (Up/Down For Delay, Escape to Exit)", Delay)
End Sub

