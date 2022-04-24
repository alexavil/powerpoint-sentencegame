Attribute VB_Name = "Module1"
Public Sub OnLoad()
    Application.ActivePresentation.Windows(1).WindowState = ppWindowMinimized
    ActivePresentation.SlideShowSettings.Run
End Sub

Sub LaunchGame()
    ActivePresentation.SlideShowWindow.View.Exit
    Presentations.Open (ActivePresentation.Path & "\to be\to_be_game.pptm")
    Application.Windows(1).Close
End Sub

Sub LaunchGame_ManyMuch()
    ActivePresentation.SlideShowWindow.View.Exit
    Presentations.Open (ActivePresentation.Path & "\many - much\many_much.pptm")
    Application.Windows(2).Close
End Sub

Public Sub ExitGame()
    Application.Quit
End Sub
