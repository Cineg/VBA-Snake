Attribute VB_Name = "ModKeys"
Option Explicit

Public Function BindKeys()
    Application.OnKey "{LEFT}", "MoveLeft"
    Application.OnKey "{RIGHT}", "MoveRight"
    Application.OnKey "{UP}", "MoveUp"
    Application.OnKey "{DOWN}", "MoveDown"
End Function

Public Function UnBindKeys()
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
End Function

Private Sub MoveLeft()
    Call ModGame.UpdateMoveDirection(0, -1)
End Sub

Private Sub MoveRight()
    Call ModGame.UpdateMoveDirection(0, 1)
End Sub

Private Sub MoveUp()
    Call ModGame.UpdateMoveDirection(-1, 0)
End Sub

Private Sub MoveDown()
    Call ModGame.UpdateMoveDirection(1, 0)
End Sub
