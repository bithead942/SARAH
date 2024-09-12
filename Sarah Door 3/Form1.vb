Public Class frmMain
    Dim WithEvents Recognizer As New SpeechRecognitionEngine
    Dim Synthesizer As New SpeechSynthesizer
    Dim WavPlay1 As New System.Media.SoundPlayer
    Dim strOccupantName As String


    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Recognizer.RecognizeAsyncCancel()

    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitializeSpeechRecognitionEngine()
        lblStatus.Text = "Sleeping"

    End Sub

    Private Sub InitializeSpeechRecognitionEngine()
        Dim customGrammar As Grammar
        Dim gBuilder As New GrammarBuilder
        Dim myCommands As New Choices

        myCommands.Add("Sydney")
        myCommands.Add("Ethan")
        myCommands.Add("Julie")
        myCommands.Add("Greg")
        myCommands.Add("No one")
        myCommands.Add("Anyone")
        myCommands.Add("I don't know")
        gBuilder.Append(myCommands)

        customGrammar = New Grammar(gBuilder)
        Recognizer.UnloadAllGrammars()
        Recognizer.LoadGrammar(customGrammar)

    End Sub

    Private Sub btnOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOn.Click
        btnOff.Enabled = True
        btnOn.Enabled = False
        lblStatus.Text = "Sleeping"
    End Sub

    Private Sub btnOff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOff.Click
        Recognizer.RecognizeAsyncCancel()
        btnOn.Enabled = True
        btnOff.Enabled = False
        lblStatus.Text = "Muted"

    End Sub

    Private Sub btnDoorbell_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoorbell.Click
        WavPlay1.SoundLocation() = "doorbell.wav"
        Try
            WavPlay1.PlaySync()
        Catch ex As Exception
            'Do nothing
        End Try

        'Play doorbell sound
        AnswerDoor()

    End Sub

    Sub AnswerDoor()
        Dim strOccupantName As String
        Dim bNameProvided As Boolean

        strOccupantName = ""
        lblStatus.Text = "Listening"
        GetVisitorName(bNameProvided)
        If bNameProvided Then
            GetOccupantName(strOccupantName)
        End If

    End Sub

    Sub GetVisitorName(ByVal bNameProvided As Boolean)
        bNameProvided = False
        Try
            Synthesizer.Speak("Hello, please state your name or your company name.")
        Catch EX As Exception
            MessageBox.Show(EX.Message)
        End Try

        lblCommand.Text = "Doorbell Rang"

        'Record voice for 5 seconds
        Recognizer.SetInputToWaveFile("VisitorName.wav")

        'Timeout
        'if voice detected, bNameProvided = True

    End Sub

    Sub GetOccupantName(ByVal strOccupantName As String)
        Try
            Synthesizer.Speak("Thank you.  Who are you here to see?")
        Catch EX As Exception
            MessageBox.Show(EX.Message)
        End Try

        lblCommand.Text = "Visitor Name Provided"
        Recognizer.SetInputToDefaultAudioDevice()
        Recognizer.RecognizeAsync(RecognizeMode.Multiple)
        'Timeout
    End Sub

    Private Sub Recognizer_SpeechRecognized(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs) Handles Recognizer.SpeechRecognized

        lblConfidence.Text = Str(e.Result.Confidence * 100)
        lblDuration.Text = e.Result.Audio.Duration.TotalMilliseconds.ToString
        lblLastInteraction.Text = Now.ToString
        strOccupantName = e.Result.Text.ToString

        If strOccupantName = "Sydney" Or strOccupantName = "Ethan" Or strOccupantName = "Greg" Or strOccupantName = "Julie" Then
            GetMessage(strOccupantName)
        End If


        'IdentifySpeaker
    End Sub


    Sub GetMessage(ByVal strOccupantName As String)
        Try
            Synthesizer.Speak(strOccupantName & "is not currently home.  Please speak a message.")
        Catch EX As Exception
            MessageBox.Show(EX.Message)
        End Try

        'Record message for 30 seconds
        Recognizer.SetInputToWaveFile("Message.wav")
        'Timeout


    End Sub


End Class
