Attribute VB_Name = "Mod_NewChatProgram"
Global ttt1 As String, ttt2 As String
Global player1TTT As Boolean, player2TTT As Boolean, userX As Boolean, played As Boolean
Global spots(0 To 9) As Boolean

Public Const original = "Tic-Tac-Toe  - "
Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
