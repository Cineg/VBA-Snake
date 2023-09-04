Attribute VB_Name = "ModGame"
Option Explicit

Type Coordinate
    row As Integer
    column As Integer
End Type

Const BOARD_SIZE As Long = 30

Private FoodPosition As Coordinate
Private Snake() As Coordinate
Private MoveDirection As Coordinate

Public Function InitializeGame() As Boolean
    On Error GoTo ErrorHandler
    Call FormatBoard
    Call CreateSnake
    FoodPosition = UpdateFoodPosition()
    
    Call ModGameClock.GameLoop
    
    InitializeGame = True
    Exit Function
ErrorHandler:
    InitializeGame = False
End Function

Public Sub FinalizeGame()
    Call ModGameClock.StopGame
    Call FormatBoard
    Call UpdateMoveDirection(0, 0)
    'Call UnBindKeys
End Sub

Public Sub UpdateMoveDirection(lngRow As Long, lngColumn As Long)
    'check valid move direction options
    If MoveDirection.row = -lngRow And MoveDirection.row <> 0 Then
        MoveDirection.row = MoveDirection.row
        Exit Sub
    End If
    
    If MoveDirection.column = -lngColumn And MoveDirection.column <> 0 Then
        MoveDirection.column = MoveDirection.column
        Exit Sub
    End If
    
    MoveDirection.row = lngRow
    MoveDirection.column = lngColumn
End Sub

Public Sub GameLoop()
    
    If CanContinue = False Then
        ModGameClock.StopGame
        MsgBox "You've lost"
        ModGame.FinalizeGame
        Exit Sub
    End If
    
    Call DrawBoard
    Call DrawSnake
    Call DrawFood
    If MoveDirection.column = 0 And MoveDirection.row = 0 Then Exit Sub
    Call UpdateSnake(MoveDirection)
    
End Sub

Private Function FormatBoard() As Boolean
    On Error GoTo ErrorHandler
    
    Const COLUMN_WIDTH As Double = 2.1
    Const ROW_HEIGHT As Double = 15
    
    Dim i As Long, j As Long
    
    With wshBoard
        With .Range(.Cells(1, 1), .Cells(BOARD_SIZE + 1, BOARD_SIZE + 1))
            .Columns.ColumnWidth = COLUMN_WIDTH
            .Rows.RowHeight = ROW_HEIGHT
            .Interior.color = RGB(0, 0, 0)
        End With
    End With

    FormatBoard = True

    Exit Function
ErrorHandler:
    FormatBoard = False
End Function

Private Function CreateSnake()
    Dim Coordinate As Coordinate
    
    Dim temporary_coordinate As Coordinate
    temporary_coordinate.column = GetRandom(0, 1)
    
    If temporary_coordinate.column = 1 Then
        temporary_coordinate.row = 0
    Else
        temporary_coordinate.row = 1
    End If
    
    Coordinate.column = GetRandom(5, BOARD_SIZE - 5)
    Coordinate.row = GetRandom(5, BOARD_SIZE - 5)
    
    
    ReDim Snake(2)
    Snake(0) = Coordinate
    Snake(1) = UpdateCoordinate(Coordinate, temporary_coordinate)
    Snake(2) = UpdateCoordinate(Coordinate, temporary_coordinate)

End Function

Private Function DrawSnake()
    Dim rgbColor As Long
    Dim lngMax As Long: lngMax = UBound(Snake)
    Dim i As Long
    For i = 0 To lngMax
        If i = 0 Then rgbColor = RGB(255, 0, 0)
        If i <> 0 Then rgbColor = GradientSnake(i, lngMax) 'RGB(100, 0, 0)
        
        With Snake(i)
            wshBoard.Cells(.row, .column).Interior.color = rgbColor
        End With
    Next i
End Function

Private Function GradientSnake(lngIndex As Long, lngMaxIndex As Long) As Long
    Dim startColor As Long: startColor = 100
    Dim endColor As Long: endColor = 255
    
    Dim progress As Double: progress = 1 - (lngIndex / lngMaxIndex)
    
    Dim color As Long
    color = Round(startColor + (endColor - startColor) * progress)
    
    GradientSnake = RGB(color, 0, 0)
End Function

Private Function DrawFood()
    Dim rgbColor As Long: rgbColor = RGB(0, 255, 0)
    With FoodPosition
        wshBoard.Cells(.row, .column).Interior.color = rgbColor
    End With
End Function

Private Function DrawBoard()
    Dim rgbColor As Long: rgbColor = RGB(0, 0, 0)
    With wshBoard
        .Range(.Cells(1, 1), .Cells(BOARD_SIZE + 1, BOARD_SIZE + 1)).Interior.color = rgbColor
    End With
End Function

Private Function UpdateSnake(direction As Coordinate)
    Dim i As Long
    Dim copyArr() As Coordinate: copyArr = Snake
    
    Snake(0) = UpdateCoordinate(Snake(0), direction)
    For i = 1 To UBound(Snake)
        Snake(i) = copyArr(i - 1)
    Next i
    
    If Snake(0).column = FoodPosition.column And Snake(0).row = FoodPosition.row Then
        ReDim Preserve Snake(UBound(Snake) + 1)
        Snake(UBound(Snake)) = copyArr(UBound(Snake) - 1)
        FoodPosition = UpdateFoodPosition
    End If
    
    Erase copyArr
End Function

Private Function UpdateCoordinate(Coordinate As Coordinate, direction As Coordinate) As Coordinate
    Coordinate.column = Coordinate.column + direction.column
    Coordinate.row = Coordinate.row + direction.row
    UpdateCoordinate = Coordinate
End Function

Private Function CanContinue() As Boolean
    Dim i As Long
    With Snake(0)
        If .column = 0 Or .row = 0 _
        Or .column = BOARD_SIZE + 2 Or .row = BOARD_SIZE + 2 Then
            CanContinue = False
            Exit Function
        End If
        
        For i = 1 To UBound(Snake)
            If .column = Snake(i).column And .row = Snake(i).row Then
                CanContinue = False
                Exit Function
            End If
        Next i
    End With
    
    CanContinue = True
End Function

Private Function UpdateFoodPosition() As Coordinate
    Dim lngRow As Long, lngColumn As Long
    Dim i As Long
    Dim Coordinate As Coordinate
        
    Dim isPossible As Boolean: isPossible = False
    
    Do While isPossible = False
        isPossible = True
        lngRow = GetRandom(1, BOARD_SIZE + 1)
        lngColumn = GetRandom(1, BOARD_SIZE + 1)
        
        For i = 0 To UBound(Snake)
            With Snake(i)
                If lngRow = .row And lngColumn = .column Then isPossible = False
                Exit For
            End With
        Next i
    Loop
    
    Coordinate.row = lngRow
    Coordinate.column = lngColumn
    UpdateFoodPosition = Coordinate
End Function

Private Function GetRandom(lngLowerBound As Long, lngUpperBound As Long) As Long
     GetRandom = Int((lngUpperBound - lngLowerBound + 1) * Rnd + lngLowerBound)
End Function
