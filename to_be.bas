Attribute VB_Name = "Module1"
#If VBA7 Then
    Declare PtrSafe Function WaitMessage Lib "user32" () As Long
    Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#Else
    Declare Function WaitMessage Lib "user32" () As Long
    Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If
Public Sentences As Integer
Public Answers As Integer
Public Correct As Integer
Public SentenceNo As Integer

Public Sub OnLoad()
For Each Slide In ActivePresentation.Slides
    For Each sh In Slide.Shapes
        If sh.Name = "Score" Then sh.Delete
    Next sh
Next Slide
CheckSentences
End Sub

Public Sub CheckSentences()
Set fs = CreateObject("Scripting.FileSystemObject")
File = fs.FileExists(ActivePresentation.Path & "\sentences.txt")
AnswerFile = fs.FileExists(ActivePresentation.Path & "\answers.txt")
If File = False Then
    SentencesMissing = "Your sentences file is missing. Would you like to load sample data?"
    Response = MsgBox(SentencesMissing, vbQuestion + vbYesNo, "Sentences File Missing")
    If Response = vbYes Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\sentences.txt")
        Sample = Array("He ... in the garden.", "They ... in the kitchen.", "We ... pupils.", "I ... in the bathroom.")
        For Each Sentence In Sample
        NewFile.WriteLine (Sentence)
        Next Sentence
        Ok = MsgBox("You can edit the sentences.txt file to add your own sentences. The sample data is loaded for you.", vbInformation, "Sentences File Created")
        CheckAnswers
    End If
    If Response = vbNo Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\sentences.txt")
        Ok = MsgBox("You can edit the sentences.txt file to add your own sentences.", vbInformation, "Sentences File Created")
    End If
End If
If File = True Then
    Sentences = CreateObject("Scripting.FileSystemObject").OpenTextFile(ActivePresentation.Path & "\sentences.txt", 8, True).Line
    Ok = MsgBox("Found " & Sentences & " sentences!", vbInformation, "Sentence Check Complete")
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
        Sample = Array("He is in the garden.", "They are in the kitchen.", "We are pupils.", "I am in the bathroom.")
        For Each Sentence In Sample
            NewFile.WriteLine (Sentence)
        Next Sentence
        Ok = MsgBox("You can edit the answers.txt file to set the answers. They must match the sentences.txt file!" & vbCrLf & "The sample data is loaded for you.", vbInformation, "Answers File Created")
    End If
        If Response = vbNo Then
        Set NewFile = fs.CreateTextFile(ActivePresentation.Path & "\answers.txt")
        Ok = MsgBox("You can edit the answers.txt file to set the answers. They must match the sentences.txt file!", vbInformation, "Answers File Created")
    End If
End If
If File = True Then
    Answers = CreateObject("Scripting.FileSystemObject").OpenTextFile(ActivePresentation.Path & "\answers.txt", 8, True).Line
    If Answers < Sentences Then Ok = MsgBox("Some of your sentences lack correct answers. Please see your answers.txt file.", vbInformation, "Answer Check Complete")
    Application.ActivePresentation.Windows(1).WindowState = ppWindowMinimized
    ActivePresentation.SlideShowSettings.Run
End If
End Sub

Public Sub DealCards()
DealCard1
DealCard2
DealCard3
End Sub

Public Sub StartGame()
Set fs = CreateObject("Scripting.FileSystemObject")
ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1)
Wait (5)
Dim rng As Integer
rng = Int((3 * Rnd) + 1)
Set Sentence = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=372, Height:=72)
Sentence.Name = "Sentence"
Set File = fs.OpenTextFile(ActivePresentation.Path & "\sentences.txt", ForReading)
Text = File.ReadLine
Sentence.TextFrame.TextRange.Text = Text
DealCard1
DealCard2
DealCard3
Correct = 0
SentenceNo = 1
End Sub

Public Sub GenerateSentence()
Counter = 1
SentenceNo = SentenceNo + 1
Set fs = CreateObject("Scripting.FileSystemObject")
If SentenceNo > Sentences Then EndGame
Dim rng As Integer
rng = Int((3 * Rnd) + 1)
Set Sentence = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=372, Height:=72)
Sentence.Name = "Sentence"
Set File = fs.OpenTextFile(ActivePresentation.Path & "\sentences.txt", ForReading)
Do While Counter < SentenceNo
    Counter = Counter + 1
    File.SkipLine
Loop
Text = File.ReadLine
If IsNull(Text) Then EndGame
Sentence.TextFrame.TextRange.Text = Text
End Sub

Public Sub Wait(Seconds As Double)
    Dim endtime As Double
    endtime = DateTime.Timer + Seconds
    Do
        WaitMessage
        DoEvents
    Loop While DateTime.Timer < endtime
End Sub

Public Sub ResetCards()
Set CurrentSlide = ActivePresentation.SlideShowWindow.View.Slide
ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card1").Delete
ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card2").Delete
ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card3").Delete
DealCards
End Sub

Public Sub DealCard1()
rng = Int((3 * Rnd) + 1)
Set shp = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=30, Left:=50, Width:=72, Height:=72)
shp.Name = "Card1"
shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
Select Case rng
Case 1
shp.TextFrame.TextRange.Text = "am"
Case 2
shp.TextFrame.TextRange.Text = "is"
Case 3
shp.TextFrame.TextRange.Text = "are"
End Select
End Sub

Public Sub DealCard2()
rng = Int((3 * Rnd) + 1)
Set shp2 = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=200, Left:=50, Width:=72, Height:=72)
shp2.Name = "Card2"
shp2.Fill.ForeColor.RGB = RGB(0, 255, 0)
Select Case rng
Case 1
shp2.TextFrame.TextRange.Text = "am"
Case 2
shp2.TextFrame.TextRange.Text = "is"
Case 3
shp2.TextFrame.TextRange.Text = "are"
End Select
End Sub

Public Sub DealCard3()
rng = Int((3 * Rnd) + 1)
Set shp3 = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=360, Left:=50, Width:=72, Height:=72)
shp3.Name = "Card3"
shp3.Fill.ForeColor.RGB = RGB(0, 0, 255)
Select Case rng
Case 1
shp3.TextFrame.TextRange.Text = "am"
Case 2
shp3.TextFrame.TextRange.Text = "is"
Case 3
shp3.TextFrame.TextRange.Text = "are"
End Select
End Sub

Public Sub CheckCard1()
Counter = 1
Set fs = CreateObject("Scripting.FileSystemObject")
Set File = fs.OpenTextFile(ActivePresentation.Path & "\answers.txt", ForReading)
Do While Counter < SentenceNo
    Counter = Counter + 1
    File.SkipLine
Loop
ans = File.ReadLine
Sentence = ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").TextFrame.TextRange.Text
Answer = Replace(Sentence, "...", ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card1").TextFrame.TextRange.Text)
Debug.Print (Sentence)
Debug.Print (Answer)
If Answer = ans Then
    Card1Correct
Else
    Call sndPlaySound32(ActivePresentation.Path & "\incorrect-buzzer.wav", 1)
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card1").Delete
    GenerateSentence
    DealCard1
End If
End Sub

Public Sub CheckCard2()
Counter = 1
Set fs = CreateObject("Scripting.FileSystemObject")
Set File = fs.OpenTextFile(ActivePresentation.Path & "\answers.txt", ForReading)
Do While Counter < SentenceNo
    Counter = Counter + 1
    File.SkipLine
Loop
ans = File.ReadLine
Sentence = ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").TextFrame.TextRange.Text
Answer = Replace(Sentence, "...", ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card2").TextFrame.TextRange.Text)
Debug.Print (Sentence)
Debug.Print (Answer)
If Answer = ans Then
    Card2Correct
Else
    Call sndPlaySound32(ActivePresentation.Path & "\incorrect-buzzer.wav", 1)
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card2").Delete
    GenerateSentence
    DealCard2
End If
End Sub

Public Sub CheckCard3()
Counter = 1
Set fs = CreateObject("Scripting.FileSystemObject")
Set File = fs.OpenTextFile(ActivePresentation.Path & "\answers.txt", ForReading)
Do While Counter < SentenceNo
    Counter = Counter + 1
    File.SkipLine
Loop
ans = File.ReadLine
Sentence = ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").TextFrame.TextRange.Text
Answer = Replace(Sentence, "...", ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card3").TextFrame.TextRange.Text)
Debug.Print (Sentence)
Debug.Print (Answer)
If Answer = ans Then
    Card3Correct
Else
    Call sndPlaySound32(ActivePresentation.Path & "\incorrect-buzzer.wav", 1)
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card3").Delete
    GenerateSentence
    DealCard3
End If
End Sub

Public Sub Card1Correct()
        Correct = Correct + 1
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card1").Delete
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
        Call sndPlaySound32(ActivePresentation.Path & "\answer-correct.wav", 1)
        DealCard1
        GenerateSentence
End Sub

Public Sub Card2Correct()
        Correct = Correct + 1
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card2").Delete
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
        Call sndPlaySound32(ActivePresentation.Path & "\answer-correct.wav", 1)
        DealCard2
        GenerateSentence
End Sub

Public Sub Card3Correct()
        Correct = Correct + 1
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card3").Delete
        ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
        Call sndPlaySound32(ActivePresentation.Path & "\answer-correct.wav", 1)
        DealCard3
        GenerateSentence
End Sub

Public Sub EndGame()
If Correct < (Sentences / 2) Then
    For Each sh In ActivePresentation.SlideShowWindow.View.Slide.Shapes
    If sh.Name = "Card1" Then Card1Exists = True
    If sh.Name = "Card2" Then Card2Exists = True
    If sh.Name = "Card3" Then Card3Exists = True
    If sh.Name = "Sentence" Then SentenceExists = True
    Next sh
If Card1Exists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card1").Delete
If Card2Exists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card2").Delete
If Card3Exists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card3").Delete
If SentenceExists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
    ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1)
    Set Score = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=372, Height:=72)
    Score.Name = "Score"
    Score.TextFrame.TextRange.Text = Correct & "/" & Sentences
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
Else
    For Each sh In ActivePresentation.SlideShowWindow.View.Slide.Shapes
    If sh.Name = "Card1" Then Card1Exists = True
    If sh.Name = "Card2" Then Card2Exists = True
    If sh.Name = "Card3" Then Card3Exists = True
    If sh.Name = "Sentence" Then SentenceExists = True
    Next sh
If Card1Exists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card1").Delete
If Card2Exists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card2").Delete
If Card3Exists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Card3").Delete
If SentenceExists = True Then ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
    ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 2)
    Set Score = ActivePresentation.SlideShowWindow.View.Slide.Shapes.AddShape(Type:=msoShapeRectangle, Top:=225, Left:=300, Width:=372, Height:=72)
    Score.Name = "Score"
    Score.TextFrame.TextRange.Text = Correct & "/" & Sentences
    ActivePresentation.SlideShowWindow.View.Slide.Shapes("Sentence").Delete
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
