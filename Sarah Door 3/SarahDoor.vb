Public Class frmMain

    Dim Synthesizer As New SpeechSynthesizer
    Dim WavPlay1 As New System.Media.SoundPlayer
    Dim iCount As Integer

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Event_HistoryTableAdapter.InsertQuery("9018")
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iHouseOccupied As Integer
        Event_HistoryTableAdapter.InsertQuery("9017")

        lblStatus.Text = "Listening"
        lblHouseOccupied.Text = "1"
        lblLastDoorbell.Text = ""
        mtxtFrequency.Text = "600"
        tCheckDoorbell.Interval = CInt(mtxtFrequency.Text)

        'Check House Occupied
        iHouseOccupied = Event_Current_StateTableAdapter.CheckHouseOccupied()
        If iHouseOccupied = 1 Then
            lblHouseOccupied.Text = "Yes"
        Else
            lblHouseOccupied.Text = "No"
        End If

        iCount = 0

        Thread.Sleep(3000)
        tCheckDoorbell.Start()
    End Sub

    Private Sub btnOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOn.Click
        btnOff.Enabled = True
        btnOn.Enabled = False
        btnDoorbell.Enabled = True
        lblStatus.Text = "Listening"
    End Sub

    Private Sub btnOff_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOff.Click
        btnOn.Enabled = True
        btnOff.Enabled = False
        btnDoorbell.Enabled = False
        lblStatus.Text = "Muted"

    End Sub

    Private Sub btnDoorbell_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDoorbell.Click

        WavPlay1.SoundLocation() = "doorbell.wav"
        Try
            WavPlay1.PlaySync()
        Catch ex As Exception
            'Do nothing
        End Try

        AnswerDoor()

    End Sub

    Sub AnswerDoor()
        Dim iHouseOccupied As Integer

        Try
            Event_HistoryTableAdapter.InsertQuery("5047")

            'Check House Occupied
            iHouseOccupied = Event_Current_StateTableAdapter.CheckHouseOccupied()
            If iHouseOccupied = 1 Then
                lblHouseOccupied.Text = "Yes"
            Else
                lblHouseOccupied.Text = "No"
            End If
            lblLastDoorbell.Text = Now.ToString

            If iHouseOccupied = 1 Then
                lblStatus.Text = "Doorbell Rang"
                Synthesizer.Speak("Welcome.  Someone will be with you shortly.")

            Else
                lblStatus.Text = "Doorbell Rang"
                Synthesizer.Speak("Sorry, my masters are not able to come to the door right now, please stop back later.")

            End If
            lblStatus.Text = "Listening"
        Catch ex As Exception
            'Do nothing
        End Try

    End Sub

    Private Sub tCheckDoorbell_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tCheckDoorbell.Tick
        Dim iHouseOccupied As Integer

        tCheckDoorbell.Stop()

        Try
            lblLastChecked.Text = Now.ToString

            If Event_Current_StateTableAdapter.CheckDoorbell() = 1 Then
                AnswerDoor()
            End If

            iCount += 1
            If iCount >= 60 Then
                'Check House Occupied
                iHouseOccupied = Event_Current_StateTableAdapter.CheckHouseOccupied()
                If iHouseOccupied = 1 Then
                    lblHouseOccupied.Text = "Yes"
                Else
                    lblHouseOccupied.Text = "No"
                End If
                iCount = 0
            End If
        Catch ex As Exception
            'Do nothing
        End Try

        tCheckDoorbell.Start()

    End Sub

    Private Sub mtxtFrequency_LostFocus(sender As Object, e As EventArgs) Handles mtxtFrequency.LostFocus
        tCheckDoorbell.Interval = CInt(mtxtFrequency.Text)
    End Sub

End Class
