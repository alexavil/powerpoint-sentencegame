Attribute VB_Name = "Module1"
#If VBA7 Then
    Declare PtrSafe Function WaitMessage Lib "user32" () As Long
    Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#Else
    Declare Function WaitMessage Lib "user32" () As Long
    Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If
Public Words As Integer
Public Answers As Integer
Public Correct As Integer
Public WordNo As Integer

Public Sub OnLoad()
For Each Slide In ActivePresentation.Slides
    For Each sh In Slide.Shapes
        If sh.Name = "Score" Then sh.Delete
    Next sh
Next Slide
CheckWords
End Sub

Public Sub CheckWords()
Set fs = CreateObject("Scripting.FileSystemObject")
File = fs.FileExists(ActivePresentation.Path & "\words.txt")
AnswerFile = fs.FileExists(ActivePresentation.Path & "\answers.txt")
If File = False Then
    WordsMissing = "Your words file is missing. Would you like to load sample data?"
    Response = MsgBox(WordsMissing, vbQuestion + vbYesNo, "Words File Missing")
    If Response = vbYes Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\words.txt")
        Sample = Array("apples", "butter", "eggs", "milk")
        For Each Word In Sample
        NewFile.WriteLine (Word)
        Next Word
        Ok = MsgBox("You can edit the words.txt file to add your own words. The sample data is loaded for you.", vbInformation, "Words File Created")
        CheckAnswers
    End If
    If Response = vbNo Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\words.txt")
        Ok = MsgBox("You can edit the words.txt file to add your own words.", vbInformation, "Words File Created")
    End If
End If
If File = True Then
    Words = CreateObject("Scripting.FileSystemObject").OpenTextFile(ActivePresentation.Path & "\words.txt", 8, True).Line
    Ok = MsgBox("Found " & Words & " words!", vbInformation, "Word Check Complete")
    CheckAnswers
    End If
End Sub

Public Sub CheckAnswers()
Set fs = CreateObject("Scripting.FileSystemObject")
File = fs.FileExists(ActivePresentation.Path & "\answers.txt")
If File = False Then
    Missing = "Your answers file is missing. Would you like to load sample data?"
    Response = MsgBox(Missing, vbQuestion + vbYesNo, "Answers File Missing")
    If Response = vbYes Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\answers.txt")
        Sample = Array("many apples", "much butter", "many eggs", "much milk")
        For Each Word In Sample
            NewFile.WriteLine (Word)
        Next Word
        Ok = MsgBox("You can edit the answers.txt file to set the answers. They must match the words.txt file!" & vbCrLf & "The sample data is loaded for you.", vbInformation, "Answers File Created")
    End If
        If Response = vbNo Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\answers.txt")
        Ok = MsgBox("You can edit the answers.txt file to set the answers. They must match the words.txt file!", vbInformation, "Answers File Created")
    End If
End If
If File = True Then
    Answers = CreateObject("Scripting.FileSystemObject").OpenTextFile(ActivePresentation.Path & "\answers.txt", 8, True).Line
    If Answers < Words Then Ok = MsgBox("Some of your words lack correct answers. Please see your answers.txt file.", vbInformation, "Answer Check Complete")
    CheckImages
End If
End Sub

Public Sub CheckImages()
Set fs = CreateObject("Scripting.FileSystemObject")
Folder = fs.FolderExists(ActivePresentation.Path & "\images")
If Folder = False Or Dir(ActivePresentation.Path & "\images" & "\*.*") = "" Then
    Alert = MsgBox("There are no images. Please insert some to get started with the game!", vbCritical, "Images Missing")
    If Folder = False Then fs.CreateFolder (ActivePresentation.Path & "\images")
End If
If Folder = True And Not (Dir(ActivePresentation.Path & "\images" & "\*.*") = "") Then
    Application.ActivePresentation.Windows(1).WindowState = ppWindowMinimized
    ActivePresentation.SlideShowSettings.Run
End If
End Sub


Public Sub StartGame()
Set fs = CreateObject("Scripting.FileSystemObject")
ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1)
Wait (5)
Dim rng As Integer
rng = Int((3 * Rnd) + 1)
Set Word = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=144, Height:=36)
Word.Name = "Word"
Set File = fs.OpenTextFile(ActivePresentation.Path & "\words.txt", ForReading)
Text = File.ReadLine
Word.TextFrame.TextRange.Text = Text
Word.Visible = False
Set Image = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddPicture(ActivePresentation.Path & "\images\1.png", msoTrue, msoTrue, Left:=380, Top:=200)
Image.Name = "Image"
Correct = 0
WordNo = 1
End Sub

Public Sub GenerateWord()
Counter = 1
WordNo = WordNo + 1
Set fs = CreateObject("Scripting.FileSystemObject")
If WordNo > Words Then EndGame
Dim rng As Integer
Set Word = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=144, Height:=36)
Word.Name = "Word"
Set File = fs.OpenTextFile(ActivePresentation.Path & "\words.txt", ForReading)
Do While Counter < WordNo
    Counter = Counter + 1
    File.SkipLine
Loop
Text = File.ReadLine
If IsNull(Text) Then EndGame
Word.TextFrame.TextRange.Text = Text
Word.Visible = False
Set Image = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddPicture(ActivePresentation.Path & "\images\" & WordNo & ".png", msoTrue, msoTrue, Left:=380, Top:=200)
Image.Name = "Image"
End Sub

Public Sub Wait(Seconds As Double)
    Dim endtime As Double
    endtime = DateTime.Timer + Seconds
    Do
        WaitMessage
        DoEvents
    Loop While DateTime.Timer < endtime
End Sub


Public Sub CheckBin1()
Counter = 1
Set fs = CreateObject("Scripting.FileSystemObject")
Set File = fs.OpenTextFile(ActivePresentation.Path & "\answers.txt", ForReading)
Do While Counter < WordNo
    Counter = Counter + 1
    File.SkipLine
Loop
ans = File.ReadLine
Word = ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").TextFrame.TextRange.Text
Answer = "many " & Word
Debug.Print (Word)
Debug.Print (Answer)
If Answer = ans Then
    IfCorrect
Else
    Call sndPlaySound32(ActivePresentation.Path & "\incorrect-buzzer.wav", 1)
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Image").Delete
    GenerateWord
End If
End Sub

Public Sub CheckBin2()
Counter = 1
Set fs = CreateObject("Scripting.FileSystemObject")
Set File = fs.OpenTextFile(ActivePresentation.Path & "\answers.txt", ForReading)
Do While Counter < WordNo
    Counter = Counter + 1
    File.SkipLine
Loop
ans = File.ReadLine
Word = ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").TextFrame.TextRange.Text
Answer = "much " & Word
Debug.Print (Word)
Debug.Print (Answer)
If Answer = ans Then
    IfCorrect
Else
    Call sndPlaySound32(ActivePresentation.Path & "\incorrect-buzzer.wav", 1)
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Image").Delete
    GenerateWord
End If
End Sub

Public Sub IfCorrect()
        Correct = Correct + 1
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Image").Delete
        Call sndPlaySound32(ActivePresentation.Path & "\answer-correct.wav", 1)
        GenerateWord
End Sub

Public Sub EndGame()
If Correct < (Words / 2) Then
    For Each sh In ActivePresentation.SlideShowWindow.View.Slide.Shapes
    If sh.Name = "Word" Then WordExists = True
    If sh.Name = "Image" Then ImageExists = True
    Next sh
If WordExists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
If ImageExists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Image").Delete
    ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1)
    Set Score = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=372, Height:=72)
    Score.Name = "Score"
    Score.TextFrame.TextRange.Text = Correct & "/" & Words
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
Else
    For Each sh In ActivePresentation.SlideShowWindow.View.Slide.Shapes
    If sh.Name = "Word" Then WordExists = True
    If sh.Name = "Image" Then ImageExists = True
    Next sh
    If WordExists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
    If ImageExists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Image").Delete
    ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 2)
    Set Score = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=372, Height:=72)
    Score.Name = "Score"
    Score.TextFrame.TextRange.Text = Correct & "/" & Words
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Word").Delete
End If
End Sub

Public Sub ExitGame()
    For Each Slide In ActivePresentation.Slides
    For Each sh In Slide.Shapes
        If sh.Name = "Score" Then sh.Delete
    Next sh
Next Slide
    Application.Quit
End Sub

