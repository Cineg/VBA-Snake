Attribute VB_Name = "ModGameClock"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Public isRunning As Boolean

Const row As Long = 2
Const COL As Long = 35

Sub StartGame()
    isRunning = True
    'ModKeys.BindKeys
    If ModGame.InitializeGame Then
        isRunning = True
    Else
        Debug.Print "Couldn't start the game :("
        Call ModGame.FinalizeGame
    End If

End Sub

Sub StopGame()
    isRunning = False
    DoEvents
End Sub

Sub GameLoop()
    Dim i As Long
    Do While isRunning = True
        Call UpdatePosition
        ModGame.GameLoop
        DoEvents
        
        For i = 0 To 6
            Sleep 10
            DoEvents
        Next i
    Loop
End Sub

Sub UpdatePosition()
    'Bind keys doesnt work for some reason XD
    'Using workaround
    If ActiveCell.row > row Then
        Call ModGame.UpdateMoveDirection(1, 0)
    End If
    If ActiveCell.row < row Then
        Call ModGame.UpdateMoveDirection(-1, 0)
    End If
    If ActiveCell.column > COL Then
        Call ModGame.UpdateMoveDirection(0, 1)
    End If
    If ActiveCell.column < COL Then
        Call ModGame.UpdateMoveDirection(0, -1)
    End If
    
    wshBoard.Cells(row, COL).Activate
End Sub
