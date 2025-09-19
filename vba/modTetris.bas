Attribute VB_Name = "modTetris"
Option Explicit

Private m_isRunning As Boolean

Public Sub SetupTetris()
    m_isRunning = False
    Debug.Print "Tetris is ready. Call StartTetris to begin."
End Sub

Public Sub StartTetris()
    If m_isRunning Then
        Debug.Print "Tetris is already running."
        Exit Sub
    End If

    m_isRunning = True
    Debug.Print "Tetris started."
    ' TODO: Implement the actual game loop and rendering logic.
End Sub

Public Sub StopTetris()
    If Not m_isRunning Then
        Debug.Print "Tetris is not running."
        Exit Sub
    End If

    m_isRunning = False
    Debug.Print "Tetris stopped."
End Sub
