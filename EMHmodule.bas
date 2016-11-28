Attribute VB_Name = "Module1"
Sub OffBoard()
If Snake.Top < 0 Then Snake.Top = 5400
If Snake.Top > 5520 Then Snake.Top = -240

If Snake.Left < 0 Then Snake.Left = 5520
If Snake.Left > 5280 Then Snake.Left = -240
End Sub
